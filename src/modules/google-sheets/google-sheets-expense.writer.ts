import { Injectable, Logger } from '@nestjs/common';
import {
  TemplateService,
  SheetSection,
} from '../template/template.service';
import {
  CategorizedTransaction,
} from '../../common/interfaces/processing-result.interface';
import { ReviewItem } from '../../common/interfaces/review-item.interface';
import {
  ExpenseCategory,
  ReviewReason,
  TransactionSource,
} from '../../common/enums/category.enum';
import { normalizeSubheading } from '../../common/utils/string.util';
import { formatDateForSheet } from '../../common/utils/date.util';
import { GoogleSheetsAdapter } from './google-sheets.adapter';

/** Maps category → heading alias for flexible matching */
const HEADING_ALIASES: Record<string, string[]> = {
  SALARY: ['Salaries'],
  BONUS: ['Team Activities & Bonuses'],
  ACTIVITY: ['Team Activities & Bonuses'],
  OPD: ['Team OPD'],
  RENT: ['DK Rent'],
  EQUIPMENT: ['Equipments', 'Equipment'],
  VENDOR: ['Vendors'],
  DONATION: ['Donations'],
  SUBSCRIPTION: ['Subscriptions'],
  PROJECT_SHARE: ['Project Shares & Commissions'],
  IHSAN_WITHDRAW: ["Ihsan's Withdrawal", 'Ihsan Withdrawl'],
  TAX_PAYMENT: ['Tax Payments'],
};

@Injectable()
export class GoogleSheetsExpenseWriter {
  private readonly logger = new Logger(GoogleSheetsExpenseWriter.name);

  constructor(
    private readonly sheetsAdapter: GoogleSheetsAdapter,
    private readonly templateService: TemplateService,
  ) {}

  /**
   * Write transactions directly to Google Sheets using insertRowsAt.
   * Preserves template structure (merged cells, formatting, spacing).
   * Does NOT reconstruct the sheet from arrays.
   */
  async writeTransactions(
    spreadsheetId: string,
    sheetName: string,
    transactions: CategorizedTransaction[],
    options: {
      templateSheetName?: string;
      overwrite?: boolean;
    } = {},
  ): Promise<{ additionalReviewItems: ReviewItem[] }> {
    const templateName = options.templateSheetName || 'Template';
    const additionalReviewItems: ReviewItem[] = [];

    // Ensure sheet exists (duplicate from Template)
    const exists = await this.sheetsAdapter.sheetExists(spreadsheetId, sheetName);
    if (!exists) {
      await this.sheetsAdapter.duplicateSheet(
        spreadsheetId,
        templateName,
        sheetName,
      );
    } else if (options.overwrite) {
      await this.sheetsAdapter.deleteSheet(spreadsheetId, sheetName);
      await this.sheetsAdapter.duplicateSheet(
        spreadsheetId,
        templateName,
        sheetName,
      );
    }

    // Read the sheet to parse structure (no XLSX rebuild)
    const rawRows = await this.sheetsAdapter.readSheet(spreadsheetId, sheetName);
    const sections = this.templateService.parseSheetSectionsFromRows(
      rawRows,
      sheetName,
    );

    // Update title cell (A1) only — preserves rest of sheet
    const monthMatch = sheetName.match(/^([A-Z]{3})-(\d{4})/i);
    if (monthMatch) {
      await this.sheetsAdapter.updateRange(
        spreadsheetId,
        sheetName,
        'A1',
        [[`MONTHLY EXPENSES - ${monthMatch[1]} ${monthMatch[2]}`]],
      );
    }

    const grouped = this.groupByCategory(transactions);
    const sortedCategories = [...grouped.entries()].sort((a, b) => {
      const secA = this.findSection(sections, a[0]);
      const secB = this.findSection(sections, b[0]);
      return (secA?.headingRow ?? Infinity) - (secB?.headingRow ?? Infinity);
    });

    let totalInserted = 0;

    for (const [category, txns] of sortedCategories) {
      const section = this.findSection(sections, category);
      if (!section) {
        this.logger.warn(
          `No section found for category "${category}", ${txns.length} transactions will be skipped`,
        );
        continue;
      }

      if (category === ExpenseCategory.SALARY) {
        const result = await this.insertSalaryRowsViaApi(
          spreadsheetId,
          sheetName,
          rawRows,
          section,
          txns,
          totalInserted,
          additionalReviewItems,
        );
        totalInserted = result.totalInserted;
      } else if (category === ExpenseCategory.SUBSCRIPTION) {
        totalInserted = await this.insertSubscriptionRowsViaApi(
          spreadsheetId,
          sheetName,
          section,
          txns,
          totalInserted,
        );
      } else {
        totalInserted = await this.insertGenericRowsViaApi(
          spreadsheetId,
          sheetName,
          section,
          txns,
          totalInserted,
        );
      }
    }

    this.logger.log(
      `Written ${transactions.length} transactions to "${sheetName}" (structure preserved)`,
    );
    return { additionalReviewItems };
  }

  private groupByCategory(
    transactions: CategorizedTransaction[],
  ): Map<string, CategorizedTransaction[]> {
    const grouped = new Map<string, CategorizedTransaction[]>();
    for (const txn of transactions) {
      const key = txn.category;
      if (!grouped.has(key)) grouped.set(key, []);
      grouped.get(key)!.push(txn);
    }
    return grouped;
  }

  private findSection(
    sections: SheetSection[],
    category: string,
  ): SheetSection | undefined {
    const aliases = HEADING_ALIASES[category] || [];
    for (const alias of aliases) {
      const sec = this.templateService.findSectionByHeading(sections, alias);
      if (sec) return sec;
    }
    return undefined;
  }

  private async insertSalaryRowsViaApi(
    spreadsheetId: string,
    sheetName: string,
    rows: unknown[][],
    section: SheetSection,
    transactions: CategorizedTransaction[],
    currentInserted: number,
    additionalReviewItems: ReviewItem[],
  ): Promise<{ totalInserted: number }> {
    const bySubheading = new Map<string, CategorizedTransaction[]>();
    const noSubheading: CategorizedTransaction[] = [];

    for (const txn of transactions) {
      if (txn.subheading) {
        if (!bySubheading.has(txn.subheading)) bySubheading.set(txn.subheading, []);
        bySubheading.get(txn.subheading)!.push(txn);
      } else {
        noSubheading.push(txn);
      }
    }

    const sectionsBottomToTop = [...(section.subheadings || [])].sort(
      (a, b) => b.row - a.row,
    );

    let totalInserted = currentInserted;

    // 1. CEO and unknown
    const ceoAndUnknown = [...bySubheading.entries()].filter(([key]) => {
      const sub = this.templateService.findSubheading(section, key);
      return !sub;
    });
    for (const [, txns] of ceoAndUnknown) {
      const totalRow = this.findTotalRow(rows, section);
      const insertAt = totalRow + totalInserted;
      const data = txns.map((t) => this.buildSalaryRow(t));
      await this.sheetsAdapter.insertRowsAt(spreadsheetId, sheetName, insertAt, data);
      totalInserted += txns.length;
    }

    // 2. No subheading
    if (noSubheading.length > 0) {
      const totalRow = this.findTotalRow(rows, section);
      const insertAt = totalRow + totalInserted;
      const data = noSubheading.map((t) => this.buildSalaryRow(t));
      await this.sheetsAdapter.insertRowsAt(spreadsheetId, sheetName, insertAt, data);
      totalInserted += noSubheading.length;
    }

    // 3. Insert under each subheading, BOTTOM TO TOP
    for (const sub of sectionsBottomToTop) {
      const key = this.findEmployeeSubheadingKey(section, sub, bySubheading);
      const txns = key ? bySubheading.get(key) ?? [] : [];
      if (txns.length === 0) continue;

      const headingRow = this.findHeadingRow(rows, sub.label, section);
      if (headingRow === -1) {
        this.logger.warn(
          `Could not find heading "${sub.label}" in sheet, skipping ${txns.length} rows → Review Required`,
        );
        for (const t of txns) {
          additionalReviewItems.push({
            date: t.date,
            beneficiary: t.beneficiaryName,
            amount: t.amount,
            currency: t.currency,
            notes: `subheading: ${sub.label}`,
            reason: ReviewReason.HEADING_NOT_FOUND,
            sourceFile: t.source as TransactionSource,
          });
        }
        continue;
      }

      const insertAt = headingRow + 1;
      const data = txns.map((t) => this.buildSalaryRow(t));
      await this.sheetsAdapter.insertRowsAt(spreadsheetId, sheetName, insertAt, data);
      totalInserted += txns.length;
    }

    return { totalInserted };
  }

  private findHeadingRow(
    rows: unknown[][],
    label: string,
    section: SheetSection,
  ): number {
    const searchNorm = normalizeSubheading(label.replace(/\s*\(\d+\)\s*$/, ''));
    const endRow = Math.min(section.totalRow + 50, rows.length);

    for (let i = section.headerRow + 1; i < endRow; i++) {
      const r = rows[i] as unknown[] | undefined;
      if (!r) continue;
      const maxCol = Math.min((r?.length ?? 0) + 10, 20);
      for (let c = 0; c < maxCol; c++) {
        const cellVal = r[c] != null ? String(r[c]).trim() : '';
        const cellNorm = normalizeSubheading(cellVal.replace(/\s*\(\d+\)\s*$/, ''));
        if (cellNorm && cellNorm === searchNorm) return i;
      }
    }
    return -1;
  }

  private findTotalRow(rows: unknown[][], section: SheetSection): number {
    for (let i = section.headerRow + 1; i < Math.min(rows.length, section.headerRow + 500); i++) {
      const r = rows[i] as unknown[] | undefined;
      if (!r) continue;
      const maxCol = Math.min((r?.length ?? 0) + 10, 20);
      for (let c = 0; c < maxCol; c++) {
        const cell = r[c] != null ? String(r[c]).trim().toUpperCase() : '';
        if (cell === 'TOTAL') return i;
      }
    }
    return section.totalRow;
  }

  private findEmployeeSubheadingKey(
    section: SheetSection,
    sub: { label: string },
    bySubheading: Map<string, CategorizedTransaction[]>,
  ): string | null {
    for (const [key] of bySubheading) {
      const found = this.templateService.findSubheading(section, key);
      if (found?.label === sub.label) return key;
    }
    return null;
  }

  private async insertSubscriptionRowsViaApi(
    spreadsheetId: string,
    sheetName: string,
    section: SheetSection,
    transactions: CategorizedTransaction[],
    offset: number,
  ): Promise<number> {
    const insertAt = section.totalRow + offset;
    const fmtDate = (d: Date) =>
      d instanceof Date ? formatDateForSheet(d) : formatDateForSheet(new Date(d));
    const data = transactions.map((txn) => [
      txn.date ? fmtDate(txn.date) : '',
      txn.description || txn.beneficiaryName,
      txn.amount,
      null,
      null,
      txn.amount,
    ]);
    await this.sheetsAdapter.insertRowsAt(spreadsheetId, sheetName, insertAt, data);
    return offset + transactions.length;
  }

  private async insertGenericRowsViaApi(
    spreadsheetId: string,
    sheetName: string,
    section: SheetSection,
    transactions: CategorizedTransaction[],
    offset: number,
  ): Promise<number> {
    const insertAt = section.totalRow + offset;
    const fmtDate = (d: Date) =>
      d instanceof Date ? formatDateForSheet(d) : formatDateForSheet(new Date(d));
    const data = transactions.map((txn) => [
      txn.date ? fmtDate(txn.date) : '',
      txn.canonicalName || txn.beneficiaryName,
      txn.description || null,
      txn.amount,
      txn.taxAmount,
      txn.grossTotal,
    ]);
    await this.sheetsAdapter.insertRowsAt(spreadsheetId, sheetName, insertAt, data);
    return offset + transactions.length;
  }

  private buildSalaryRow(txn: CategorizedTransaction): unknown[] {
    const dateVal =
      txn.date instanceof Date
        ? formatDateForSheet(txn.date)
        : txn.date
          ? formatDateForSheet(new Date(txn.date))
          : '';
    return [
      dateVal,
      txn.canonicalName || txn.beneficiaryName,
      txn.designation || '',
      txn.amount,
      txn.taxAmount,
      txn.grossTotal,
      txn.bonus,
      txn.cnic || '',
    ];
  }
}
