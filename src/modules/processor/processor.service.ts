import { Injectable, Logger } from '@nestjs/common';
import * as XLSX from 'xlsx';
import { ConfigService } from '../config/config.service';
import { ExcelService } from '../excel/excel.service';
import { ScBankParser } from '../excel/parsers/sc-bank.parser';
import { PayoneerParser } from '../excel/parsers/payoneer.parser';
import { CategorizerService } from './categorizer.service';
import { EmployeeMatcherService } from './employee-matcher.service';
import { TemplateService } from '../template/template.service';
import { ExpenseSheetWriter } from '../excel/writers/expense-sheet.writer';
import {
  CategorizedTransaction,
  ProcessingResult,
  ProcessingStats,
} from '../../common/interfaces/processing-result.interface';
import { ReviewItem } from '../../common/interfaces/review-item.interface';
import {
  ReviewReason,
  TransactionSource,
} from '../../common/enums/category.enum';
import { normalizeSubheading } from '../../common/utils/string.util';
import { formatMonthLabel } from '../../common/utils/date.util';

export interface ProcessOptions {
  configFilePath: string;
  scStatementPath: string;
  payoneerStatementPath?: string;
  expensesWorkbookPath: string;
  outputPath: string;
  targetMonth?: string;
  templateSheetName?: string;
  overwrite?: boolean;
}

@Injectable()
export class ProcessorService {
  private readonly logger = new Logger(ProcessorService.name);

  constructor(
    private readonly configService: ConfigService,
    private readonly excelService: ExcelService,
    private readonly scBankParser: ScBankParser,
    private readonly payoneerParser: PayoneerParser,
    private readonly categorizer: CategorizerService,
    private readonly employeeMatcher: EmployeeMatcherService,
    private readonly templateService: TemplateService,
    private readonly sheetWriter: ExpenseSheetWriter,
  ) {}

  async process(options: ProcessOptions): Promise<ProcessingResult> {
    this.logger.log('=== Starting Expense Automation Processing ===');

    // Step 1: Load configuration
    this.logger.log('Step 1: Loading configuration...');
    this.configService.loadFromWorkbook(options.configFilePath);

    // Step 2: Parse SC Bank Statement
    this.logger.log('Step 2: Parsing SC Bank Statement...');
    const scTransactions = this.scBankParser.parse(options.scStatementPath);

    // Step 3: Parse Payoneer Statement (optional)
    let payoneerTransactions: ReturnType<PayoneerParser['parse']> = [];
    if (options.payoneerStatementPath) {
      this.logger.log('Step 3: Parsing Payoneer Statement...');
      payoneerTransactions = this.payoneerParser.parse(
        options.payoneerStatementPath,
      );
    } else {
      this.logger.log('Step 3: No Payoneer statement provided, skipping');
    }

    // Step 4: Determine target month
    const monthLabel =
      options.targetMonth ||
      this.inferMonth(scTransactions.map((t) => t.paymentDate));
    this.logger.log(`Target month: ${monthLabel}`);

    // Step 5: Categorize all transactions
    this.logger.log('Step 5: Categorizing transactions...');
    const categorized: CategorizedTransaction[] = [];
    const reviewItems: ReviewItem[] = [];
    const salarySeen = new Map<string, boolean>();

    for (const txn of scTransactions) {
      const result = this.categorizer.categorizeScTransaction(txn);

      if (result.review) {
        reviewItems.push(result.review);
        continue;
      }

      if (result.categorized) {
        // Duplicate salary check
        const rule = this.configService.getCategoryRule(
          result.categorized.category,
        );
        if (rule && rule.isSalary && !rule.allowMultiplePerMonth) {
          const key = `${result.categorized.employeeId}_${result.categorized.category}`;
          if (salarySeen.has(key)) {
            reviewItems.push({
              date: txn.paymentDate,
              beneficiary: txn.beneficiaryName,
              amount: txn.amount,
              currency: txn.currency,
              notes: txn.notesToSelf,
              reason: ReviewReason.DUPLICATE_SALARY,
              sourceFile: TransactionSource.SC_BANK,
            });
            continue;
          }
          salarySeen.set(key, true);
        }

        categorized.push(result.categorized);
      }
    }

    // Categorize Payoneer transactions
    for (const txn of payoneerTransactions) {
      categorized.push(this.categorizer.categorizePayoneerTransaction(txn));
    }

    this.logger.log(
      `Categorized: ${categorized.length}, Review: ${reviewItems.length}`,
    );

    // Step 6: Prepare expenses workbook
    this.logger.log('Step 6: Preparing expenses workbook...');
    const expensesWorkbook = this.excelService.readWorkbook(
      options.expensesWorkbookPath,
    );

    const templateSheet = options.templateSheetName || 'Template';
    const actualSheetName = this.templateService.ensureMonthlySheet(
      expensesWorkbook,
      monthLabel,
      templateSheet,
      options.overwrite ?? false,
    );

    // Step 6b: Filter salary transactions with invalid subheadings → Review Required
    const sections = this.templateService.parseSheetSections(
      expensesWorkbook,
      actualSheetName,
    );
    const validSubheadings = this.templateService.getValidSalarySubheadings(
      sections,
    );
    const { filteredCategorized, subheadingReviewItems } =
      this.filterByValidSubheadings(categorized, validSubheadings);
    reviewItems.push(...subheadingReviewItems);

    // Step 7: Write categorized transactions
    this.logger.log(
      `Step 7: Writing transactions to "${actualSheetName}"...`,
    );
    const writeResult = this.sheetWriter.writeTransactions(
      expensesWorkbook,
      actualSheetName,
      filteredCategorized,
    );
    reviewItems.push(...writeResult.additionalReviewItems);

    // Step 8: Write Review Required sheet
    this.logger.log('Step 8: Writing Review Required sheet...');
    this.sheetWriter.writeReviewSheet(expensesWorkbook, reviewItems);

    // Step 9: Save output
    this.logger.log('Step 9: Saving output...');
    this.excelService.writeWorkbook(expensesWorkbook, options.outputPath);

    // Build stats
    const stats = this.buildStats(
      scTransactions.length,
      payoneerTransactions.length,
      filteredCategorized,
      reviewItems,
    );

    this.logger.log('=== Processing Complete ===');
    this.printSummary(stats, actualSheetName, options.outputPath);

    return {
      monthLabel: actualSheetName,
      categorized: filteredCategorized,
      reviewItems,
      stats,
    };
  }

  /**
   * Filter salary transactions: those with subheading not in validSubheadings
   * go to Review Required with INVALID_SUBHEADING.
   */
  private filterByValidSubheadings(
    categorized: CategorizedTransaction[],
    validSubheadings: Set<string>,
  ): {
    filteredCategorized: CategorizedTransaction[];
    subheadingReviewItems: ReviewItem[];
  } {
    const filtered: CategorizedTransaction[] = [];
    const reviewItems: ReviewItem[] = [];

    for (const txn of categorized) {
      if (txn.category !== 'SALARY') {
        filtered.push(txn);
        continue;
      }
      if (!txn.subheading || txn.subheading.toUpperCase() === 'CEO') {
        filtered.push(txn);
        continue;
      }
      const norm = normalizeSubheading(txn.subheading);
      if (!validSubheadings.has(norm)) {
        reviewItems.push({
          date: txn.date,
          beneficiary: txn.beneficiaryName,
          amount: txn.amount,
          currency: txn.currency,
          notes: `salary_subheading: ${txn.subheading}`,
          reason: ReviewReason.INVALID_SUBHEADING,
          sourceFile: txn.source as TransactionSource,
        });
        this.logger.warn(
          `Salary subheading "${txn.subheading}" not in template for ${txn.canonicalName || txn.beneficiaryName} → Review Required`,
        );
      } else {
        filtered.push(txn);
      }
    }

    return { filteredCategorized: filtered, subheadingReviewItems: reviewItems };
  }

  private inferMonth(dates: Date[]): string {
    if (dates.length === 0) {
      const now = new Date();
      return formatMonthLabel(now);
    }

    // Use the most common month in the transactions
    const monthCounts = new Map<string, number>();
    for (const d of dates) {
      const label = formatMonthLabel(d);
      monthCounts.set(label, (monthCounts.get(label) || 0) + 1);
    }

    let maxLabel = '';
    let maxCount = 0;
    for (const [label, count] of monthCounts) {
      if (count > maxCount) {
        maxCount = count;
        maxLabel = label;
      }
    }

    return maxLabel;
  }

  private buildStats(
    totalSc: number,
    totalPayoneer: number,
    categorized: CategorizedTransaction[],
    reviewItems: ReviewItem[],
  ): ProcessingStats {
    const byCategory: Record<string, number> = {};
    for (const txn of categorized) {
      byCategory[txn.category] = (byCategory[txn.category] || 0) + 1;
    }

    return {
      totalScTransactions: totalSc,
      totalPayoneerTransactions: totalPayoneer,
      categorizedCount: categorized.length,
      reviewCount: reviewItems.length,
      skippedCount:
        totalSc + totalPayoneer - categorized.length - reviewItems.length,
      byCategory,
    };
  }

  private printSummary(
    stats: ProcessingStats,
    sheetName: string,
    outputPath: string,
  ): void {
    this.logger.log('');
    this.logger.log('┌─────────────────────────────────────────┐');
    this.logger.log('│         PROCESSING SUMMARY              │');
    this.logger.log('├─────────────────────────────────────────┤');
    this.logger.log(
      `│  SC Transactions:     ${String(stats.totalScTransactions).padStart(6)}          │`,
    );
    this.logger.log(
      `│  Payoneer Transactions: ${String(stats.totalPayoneerTransactions).padStart(4)}          │`,
    );
    this.logger.log(
      `│  Categorized:         ${String(stats.categorizedCount).padStart(6)}          │`,
    );
    this.logger.log(
      `│  Review Required:     ${String(stats.reviewCount).padStart(6)}          │`,
    );
    this.logger.log('├─────────────────────────────────────────┤');
    this.logger.log('│  By Category:                           │');
    for (const [cat, count] of Object.entries(stats.byCategory)) {
      this.logger.log(
        `│    ${cat.padEnd(22)} ${String(count).padStart(4)}          │`,
      );
    }
    this.logger.log('├─────────────────────────────────────────┤');
    this.logger.log(`│  Output Sheet: ${sheetName.padEnd(24)} │`);
    this.logger.log(`│  Output File:  ${outputPath.slice(-24).padEnd(24)} │`);
    this.logger.log('└─────────────────────────────────────────┘');
  }
}
