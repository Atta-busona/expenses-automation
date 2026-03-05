import { Injectable, Logger } from '@nestjs/common';
import { GoogleDriveService, DriveFile } from '../google-drive/google-drive.service';
import { GoogleSheetsAdapter } from '../google-sheets/google-sheets.adapter';
import { GoogleSheetsExpenseWriter } from '../google-sheets/google-sheets-expense.writer';
import { ConfigService } from '../config/config.service';
import { ScBankParser } from '../excel/parsers/sc-bank.parser';
import { PayoneerParser } from '../excel/parsers/payoneer.parser';
import { CategorizerService } from './categorizer.service';
import { ExcelService } from '../excel/excel.service';
import { TemplateService } from '../template/template.service';
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
import {
  formatDateForSheet,
  formatMonthLabel,
} from '../../common/utils/date.util';

export interface GoogleProcessOptions {
  configSpreadsheetId: string;
  expensesSpreadsheetId: string;
  scStatementDriveFile: DriveFile;
  payoneerStatementDriveFile?: DriveFile;
  targetMonth?: string;
  templateSheetName?: string;
  overwrite?: boolean;
}

@Injectable()
export class GoogleProcessorService {
  private readonly logger = new Logger(GoogleProcessorService.name);

  constructor(
    private readonly driveService: GoogleDriveService,
    private readonly sheetsAdapter: GoogleSheetsAdapter,
    private readonly configService: ConfigService,
    private readonly scBankParser: ScBankParser,
    private readonly payoneerParser: PayoneerParser,
    private readonly categorizer: CategorizerService,
    private readonly excelService: ExcelService,
    private readonly templateService: TemplateService,
    private readonly sheetsExpenseWriter: GoogleSheetsExpenseWriter,
  ) {}

  async processFromDrive(
    options: GoogleProcessOptions,
  ): Promise<ProcessingResult> {
    this.logger.log('=== Starting Google Drive Expense Processing ===');

    // Step 1: Load config from Google Sheets
    this.logger.log('Step 1: Loading config from Google Sheets...');
    await this.loadConfigFromSheets(options.configSpreadsheetId);

    // Step 2: Download SC Bank Statement from Drive
    this.logger.log(
      `Step 2: Downloading SC statement "${options.scStatementDriveFile.name}"...`,
    );
    const scLocalPath = await this.driveService.getFileForProcessing(
      options.scStatementDriveFile,
    );
    const scTransactions = this.scBankParser.parse(scLocalPath);

    // Step 3: Download Payoneer Statement (optional)
    let payoneerTransactions: ReturnType<PayoneerParser['parse']> = [];
    if (options.payoneerStatementDriveFile) {
      this.logger.log(
        `Step 3: Downloading Payoneer statement "${options.payoneerStatementDriveFile.name}"...`,
      );
      const payLocalPath = await this.driveService.getFileForProcessing(
        options.payoneerStatementDriveFile,
      );
      payoneerTransactions = this.payoneerParser.parse(payLocalPath);
    } else {
      this.logger.log('Step 3: No Payoneer statement, skipping');
    }

    // Step 4: Determine target month
    const monthLabel =
      options.targetMonth ||
      this.inferMonth(scTransactions.map((t) => t.paymentDate));
    this.logger.log(`Target month: ${monthLabel}`);

    // Step 5: Categorize all transactions
    this.logger.log('Step 5: Categorizing transactions...');
    const { categorized, reviewItems } = this.categorizeAll(
      scTransactions,
      payoneerTransactions,
    );

    this.logger.log(
      `Categorized: ${categorized.length}, Review: ${reviewItems.length}`,
    );

    // Step 6: Determine actual sheet name (handle overwrite / duplicate)
    const templateSheet = options.templateSheetName || 'Template';
    const actualSheetName = await this.ensureMonthlySheetName(
      options.expensesSpreadsheetId,
      monthLabel,
      templateSheet,
      options.overwrite ?? false,
    );
    this.logger.log(`Step 6: Monthly sheet: "${actualSheetName}"`);

    // Step 7: Get valid subheadings from Template (read directly from Sheets)
    const templateRows = await this.sheetsAdapter.readSheet(
      options.expensesSpreadsheetId,
      templateSheet,
    );
    const sections = this.templateService.parseSheetSectionsFromRows(
      templateRows,
      templateSheet,
    );
    const validSubheadings = this.templateService.getValidSalarySubheadings(
      sections,
    );
    const { filteredCategorized, subheadingReviewItems } =
      this.filterByValidSubheadings(categorized, validSubheadings);
    reviewItems.push(...subheadingReviewItems);

    // Step 8: Write transactions directly to Google Sheets (insertRowsAt)
    // Preserves template structure: merged cells, formatting, spacing
    this.logger.log(
      `Step 8: Writing transactions to "${actualSheetName}" (structure preserved)...`,
    );
    const writeResult = await this.sheetsExpenseWriter.writeTransactions(
      options.expensesSpreadsheetId,
      actualSheetName,
      filteredCategorized,
      {
        templateSheetName: templateSheet,
        overwrite: options.overwrite,
      },
    );
    reviewItems.push(...writeResult.additionalReviewItems);

    // Step 9: Write Review Required sheet
    this.logger.log('Step 9: Writing Review Required sheet...');
    await this.writeReviewSheetToGoogle(
      options.expensesSpreadsheetId,
      reviewItems,
    );

    const stats = this.buildStats(
      scTransactions.length,
      payoneerTransactions.length,
      filteredCategorized,
      reviewItems,
    );

    this.logger.log('=== Google Drive Processing Complete ===');
    this.logger.log(`Output: "${actualSheetName}" sheet in Google Sheets`);
    this.logger.log(
      `Categorized: ${stats.categorizedCount}, Review: ${stats.reviewCount}`,
    );

    return {
      monthLabel: actualSheetName,
      categorized: filteredCategorized,
      reviewItems,
      stats,
    };
  }

  /**
   * Load CATEGORY_RULES and EMPLOYEE_MASTER directly from Google Sheets
   * (no need to download as XLSX for config data).
   */
  private async loadConfigFromSheets(
    configSpreadsheetId: string,
  ): Promise<void> {
    const categoryRows = await this.sheetsAdapter.readSheetAsJson(
      configSpreadsheetId,
      'CATEGORY_RULES',
    );
    const employeeRows = await this.sheetsAdapter.readSheetAsJson(
      configSpreadsheetId,
      'Employee_Master',
    );

    this.configService.loadFromParsedData(categoryRows, employeeRows);
  }

  /**
   * Determine the actual monthly sheet name (handle exists/overwrite).
   */
  private async ensureMonthlySheetName(
    spreadsheetId: string,
    monthLabel: string,
    templateSheetName: string,
    overwrite: boolean,
  ): Promise<string> {
    const exists = await this.sheetsAdapter.sheetExists(
      spreadsheetId,
      monthLabel,
    );
    if (!exists) return monthLabel;
    if (overwrite) return monthLabel;
    let suffix = 2;
    let altName = `${monthLabel}-${suffix}`;
    while (await this.sheetsAdapter.sheetExists(spreadsheetId, altName)) {
      suffix++;
      altName = `${monthLabel}-${suffix}`;
    }
    this.logger.log(
      `Sheet "${monthLabel}" exists, using "${altName}" instead`,
    );
    return altName;
  }

  /**
   * Write Review Required sheet to Google Sheets.
   */
  private async writeReviewSheetToGoogle(
    spreadsheetId: string,
    reviewItems: ReviewItem[],
  ): Promise<void> {
    if (reviewItems.length === 0) return;

    const fmtDate = (d: Date) =>
      d instanceof Date ? formatDateForSheet(d) : formatDateForSheet(new Date(d));
    const data: unknown[][] = [
      ['Date', 'Beneficiary', 'Amount', 'Currency', 'Notes', 'Reason', 'Source File'],
      ...reviewItems.map((item) => [
        item.date ? fmtDate(item.date) : '',
        item.beneficiary,
        item.amount,
        item.currency,
        item.notes || '',
        item.reason,
        item.sourceFile,
      ]),
    ];

    const reviewExists = await this.sheetsAdapter.sheetExists(
      spreadsheetId,
      'Review Required',
    );
    if (!reviewExists) {
      await this.sheetsAdapter.duplicateSheet(
        spreadsheetId,
        'Template',
        'Review Required',
      );
    }
    await this.sheetsAdapter.writeToSheet(
      spreadsheetId,
      'Review Required',
      data,
    );
    this.logger.log(`Written ${reviewItems.length} items to "Review Required"`);
  }

  private categorizeAll(
    scTransactions: ReturnType<ScBankParser['parse']>,
    payoneerTransactions: ReturnType<PayoneerParser['parse']>,
  ): { categorized: CategorizedTransaction[]; reviewItems: ReviewItem[] } {
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

    for (const txn of payoneerTransactions) {
      categorized.push(this.categorizer.categorizePayoneerTransaction(txn));
    }

    return { categorized, reviewItems };
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
      // CEO and no-subheading: no validation needed
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
    if (dates.length === 0) return formatMonthLabel(new Date());
    const counts = new Map<string, number>();
    for (const d of dates) {
      const label = formatMonthLabel(d);
      counts.set(label, (counts.get(label) || 0) + 1);
    }
    let best = '';
    let max = 0;
    for (const [label, count] of counts) {
      if (count > max) {
        max = count;
        best = label;
      }
    }
    return best;
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
}
