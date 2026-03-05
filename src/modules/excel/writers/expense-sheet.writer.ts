import { Injectable, Logger } from '@nestjs/common';
import * as XLSX from 'xlsx';
import { ExcelService } from '../excel.service';
import {
  TemplateService,
  SheetSection,
} from '../../template/template.service';
import {
  CategorizedTransaction,
} from '../../../common/interfaces/processing-result.interface';
import { ReviewItem } from '../../../common/interfaces/review-item.interface';
import {
  ExpenseCategory,
  ReviewReason,
  TransactionSource,
} from '../../../common/enums/category.enum';
import { normalizeSubheading } from '../../../common/utils/string.util';
import { formatDateForSheet } from '../../../common/utils/date.util';

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
export class ExpenseSheetWriter {
  private readonly logger = new Logger(ExpenseSheetWriter.name);

  constructor(
    private readonly excelService: ExcelService,
    private readonly templateService: TemplateService,
  ) {}

  writeTransactions(
    workbook: XLSX.WorkBook,
    sheetName: string,
    transactions: CategorizedTransaction[],
  ): { additionalReviewItems: ReviewItem[] } {
    const sections = this.templateService.parseSheetSections(
      workbook,
      sheetName,
    );

    const rawRows = this.excelService.sheetToRawRows(workbook, sheetName);
    const rows: unknown[][] = rawRows.map((r) => [...(r || [])]);

    // Explicitly preserve all salary subheading labels (template may lose them on XLSX export)
    this.preserveSalarySubheadingLabels(rows, sections);

    // Update title row to target month (e.g. "MONTHLY EXPENSES - FEB 2026")
    const monthMatch = sheetName.match(/^([A-Z]{3})-(\d{4})/i);
    if (monthMatch && rows[0]) {
      rows[0][0] = `MONTHLY EXPENSES - ${monthMatch[1]} ${monthMatch[2]}`;
    }

    const grouped = this.groupByCategory(transactions);

    // Sort categories by their section position in the sheet (top→bottom)
    // so offset accumulates correctly as we insert rows downward.
    const sortedCategories = [...grouped.entries()].sort((a, b) => {
      const secA = this.findSection(sections, a[0]);
      const secB = this.findSection(sections, b[0]);
      return (secA?.headingRow ?? Infinity) - (secB?.headingRow ?? Infinity);
    });

    let rowOffset = 0;
    let salaryInsertedRanges: { startRow: number; count: number; formatSourceRow: number }[] = [];
    const additionalReviewItems: ReviewItem[] = [];

    for (const [category, txns] of sortedCategories) {
      const section = this.findSection(sections, category);
      if (!section) {
        this.logger.warn(
          `No section found for category "${category}", ${txns.length} transactions will be skipped`,
        );
        continue;
      }

      if (category === ExpenseCategory.SALARY) {
        const result = this.insertSalaryRows(
          rows,
          section,
          txns,
          rowOffset,
          additionalReviewItems,
        );
        rowOffset = result.totalInserted;
        salaryInsertedRanges = result.insertedRanges;
      } else if (category === ExpenseCategory.SUBSCRIPTION) {
        rowOffset = this.insertSubscriptionRows(
          rows,
          section,
          txns,
          rowOffset,
        );
      } else {
        rowOffset = this.insertGenericRows(
          rows,
          section,
          txns,
          rowOffset,
        );
      }
    }

    // Convert rows back to worksheet
    const sourceSheet = workbook.Sheets[sheetName];
    const newSheet = XLSX.utils.aoa_to_sheet(rows);
    this.preserveColumnWidths(sourceSheet, newSheet);
    this.preserveRowFormatting(sourceSheet, newSheet, salaryInsertedRanges);
    workbook.Sheets[sheetName] = newSheet;

    this.logger.log(
      `Written ${transactions.length} transactions to "${sheetName}"`,
    );
    return { additionalReviewItems };
  }

  writeReviewSheet(
    workbook: XLSX.WorkBook,
    reviewItems: ReviewItem[],
  ): void {
    if (reviewItems.length === 0) {
      this.logger.log('No review items to write');
      return;
    }

    const headers = [
      'Date',
      'Beneficiary',
      'Amount',
      'Currency',
      'Notes',
      'Reason',
      'Source File',
    ];

    const data: unknown[][] = [headers];

    const fmtDate = (d: Date) =>
      d instanceof Date ? formatDateForSheet(d) : formatDateForSheet(new Date(d));
    for (const item of reviewItems) {
      data.push([
        item.date ? fmtDate(item.date) : '',
        item.beneficiary,
        item.amount,
        item.currency,
        item.notes || '',
        item.reason,
        item.sourceFile,
      ]);
    }

    const sheet = this.excelService.createSheet(data);

    // Apply column widths
    sheet['!cols'] = [
      { wch: 12 },
      { wch: 30 },
      { wch: 15 },
      { wch: 8 },
      { wch: 20 },
      { wch: 40 },
      { wch: 20 },
    ];

    this.excelService.addSheetToWorkbook(
      workbook,
      'Review Required',
      sheet,
    );

    this.logger.log(
      `Written ${reviewItems.length} items to "Review Required" sheet`,
    );
  }

  private groupByCategory(
    transactions: CategorizedTransaction[],
  ): Map<string, CategorizedTransaction[]> {
    const grouped = new Map<string, CategorizedTransaction[]>();
    for (const txn of transactions) {
      const key = txn.category;
      if (!grouped.has(key)) {
        grouped.set(key, []);
      }
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
      const section = this.templateService.findSectionByHeading(
        sections,
        alias,
      );
      if (section) return section;
    }

    // Fallback: try direct heading match from the transaction
    return undefined;
  }

  /**
   * Insert salary rows using bottom-to-top order. Before each insertion, we
   * SEARCH the current rows array for the heading — we do NOT rely on stored
   * row indexes, since earlier insertions shift rows and make stored indexes wrong.
   */
  private insertSalaryRows(
    rows: unknown[][],
    section: SheetSection,
    transactions: CategorizedTransaction[],
    _currentOffset: number,
    additionalReviewItems: ReviewItem[],
  ): { totalInserted: number; insertedRanges: { startRow: number; count: number; formatSourceRow: number }[] } {
    const insertedRanges: { startRow: number; count: number; formatSourceRow: number }[] = [];

    // Group salary transactions by subheading (department)
    const bySubheading = new Map<string, CategorizedTransaction[]>();
    const noSubheading: CategorizedTransaction[] = [];

    for (const txn of transactions) {
      if (txn.subheading) {
        if (!bySubheading.has(txn.subheading)) {
          bySubheading.set(txn.subheading, []);
        }
        bySubheading.get(txn.subheading)!.push(txn);
      } else {
        noSubheading.push(txn);
      }
    }

    // Subheadings sorted BOTTOM TO TOP (highest row first) — process order only
    const sectionsBottomToTop = [...(section.subheadings || [])].sort(
      (a, b) => b.row - a.row,
    );

    let totalInserted = 0;

    // 1. Insert CEO and unknown subheadings first (before Total)
    const ceoAndUnknown = [...bySubheading.entries()].filter(([key]) => {
      const sub = this.templateService.findSubheading(section, key);
      return !sub;
    });
    for (const [subheading, txns] of ceoAndUnknown) {
      const insertAt = this.findTotalRow(rows, section) + totalInserted;
      if (subheading.toUpperCase() === 'CEO') {
        this.logger.log(
          `Inserting ${txns.length} CEO salary row(s) before Total at row ${insertAt + 1}`,
        );
      } else {
        this.logger.warn(
          `Salary subheading "${subheading}" not found in template, inserting before Total`,
        );
      }
      for (let i = 0; i < txns.length; i++) {
        const row = this.buildSalaryRow(txns[i]);
        rows.splice(insertAt + i, 0, row);
        totalInserted++;
      }
      if (txns.length > 0) {
        insertedRanges.push({
          startRow: insertAt,
          count: txns.length,
          formatSourceRow: Math.max(section.headerRow + 1, section.totalRow - 1),
        });
      }
    }

    // 2. Insert noSubheading before Total
    if (noSubheading.length > 0) {
      const insertAt = this.findTotalRow(rows, section) + totalInserted;
      for (let i = 0; i < noSubheading.length; i++) {
        const row = this.buildSalaryRow(noSubheading[i]);
        rows.splice(insertAt + i, 0, row);
        totalInserted++;
      }
      insertedRanges.push({
        startRow: insertAt,
        count: noSubheading.length,
        formatSourceRow: Math.max(section.headerRow + 1, section.totalRow - 1),
      });
    }

    // 3. Insert under each subheading, BOTTOM TO TOP — SEARCH for heading each time
    for (const sub of sectionsBottomToTop) {
      const key = this.findEmployeeSubheadingKey(section, sub, bySubheading);
      const txns = key ? bySubheading.get(key) ?? [] : [];
      if (txns.length === 0) continue;

      const { row: headingRow, foundText: foundHeadingText } = this.findHeadingRow(
        rows,
        sub.label,
        section,
      );

      // Debug logging before insert
      for (const t of txns) {
        this.logger.log(
          `[Salary insert] employeeName=${t.canonicalName || t.beneficiaryName} | ` +
            `salary_subheading=${t.subheading} | ` +
            `foundHeadingText=${foundHeadingText ?? '(not found)'} | ` +
            `headingRow=${headingRow >= 0 ? headingRow + 1 : -1}`,
        );
      }

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
      this.logger.log(
        `Inserting ${txns.length} salary row(s) under "${sub.label}" at row ${insertAt + 1} (found heading at row ${headingRow + 1})`,
      );
      for (let i = 0; i < txns.length; i++) {
        const row = this.buildSalaryRow(txns[i]);
        rows.splice(insertAt + i, 0, row);
        totalInserted++;
      }
      insertedRanges.push({
        startRow: insertAt,
        count: txns.length,
        formatSourceRow: headingRow + 1,
      });
    }

    this.ensureSubheadingLabelsAfterInsertBottomToTop(rows, section);

    return { totalInserted, insertedRanges };
  }

  /**
   * Search the current rows array for a heading by label. EXACT match only.
   * Scans the ENTIRE row (all cells) to handle merged cells where text may
   * appear in any column. Returns { row, foundText } or { row: -1, foundText: null }.
   */
  private findHeadingRow(
    rows: unknown[][],
    label: string,
    section: SheetSection,
  ): { row: number; foundText: string | null } {
    const searchNorm = normalizeSubheading(label.replace(/\s*\(\d+\)\s*$/, ''));
    const endRow = Math.min(section.totalRow + 50, rows.length);

    for (let i = section.headerRow + 1; i < endRow; i++) {
      const r = rows[i];
      if (!r) continue;
      const maxCol = Math.min((r?.length ?? 0) + 10, 20);
      for (let c = 0; c < maxCol; c++) {
        const cellVal = r[c] != null ? String(r[c]).trim() : '';
        const cellNorm = normalizeSubheading(cellVal.replace(/\s*\(\d+\)\s*$/, ''));
        if (!cellNorm) continue;
        if (cellNorm === searchNorm) {
          return { row: i, foundText: cellVal || null };
        }
      }
    }
    return { row: -1, foundText: null };
  }

  /**
   * Search for the Total row in the salaries section (may have shifted).
   * Scans entire row to handle merged cells.
   */
  private findTotalRow(rows: unknown[][], section: SheetSection): number {
    for (let i = section.headerRow + 1; i < Math.min(rows.length, section.headerRow + 500); i++) {
      const r = rows[i];
      if (!r) continue;
      const maxCol = Math.min((r?.length ?? 0) + 10, 20);
      for (let c = 0; c < maxCol; c++) {
        const cell = r[c] != null ? String(r[c]).trim().toUpperCase() : '';
        if (cell === 'TOTAL') return i;
      }
    }
    return section.totalRow;
  }

  /** Find employee subheading key that matches this template subheading */
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

  /**
   * Ensure subheading labels are present after bottom-to-top insertions.
   * With bottom-to-top, subheading rows above the insertion point are never shifted,
   * so we only need to re-apply labels for subheadings that might have been empty.
   */
  private ensureSubheadingLabelsAfterInsertBottomToTop(
    rows: unknown[][],
    section: SheetSection,
  ): void {
    for (const sub of section.subheadings ?? []) {
      if (sub.row < rows.length) {
        const col = sub.col ?? 0;
        const current = rows[sub.row]?.[col];
        const expected = sub.label;
        const firstWord = expected.replace(/\s*\(\d+\)\s*$/, '').trim().split(' ')[0];
        if (
          current == null ||
          String(current).trim() === '' ||
          !String(current).trim().toUpperCase().includes(firstWord.toUpperCase())
        ) {
          if (!rows[sub.row]) rows[sub.row] = [];
          (rows[sub.row] as unknown[])[col] = sub.label;
        }
      }
    }
  }

  /**
   * Preserve salary subheading labels in rows before any insertions.
   * Template may lose text on XLSX export (merged cells, etc.).
   */
  private preserveSalarySubheadingLabels(
    rows: unknown[][],
    sections: SheetSection[],
  ): void {
    const salarySection = sections.find((s) =>
      s.headingText.toLowerCase().includes('salar'),
    );
    if (!salarySection?.subheadings?.length) return;

    for (const sub of salarySection.subheadings) {
      if (sub.row < rows.length) {
        if (!rows[sub.row]) rows[sub.row] = [];
        const col = sub.col ?? 0;
        (rows[sub.row] as unknown[])[col] = sub.label;
      }
    }
  }

  private insertSubscriptionRows(
    rows: unknown[][],
    section: SheetSection,
    transactions: CategorizedTransaction[],
    currentOffset: number,
  ): number {
    let offset = currentOffset;
    const insertAt = section.totalRow + offset;

    const fmtDate = (d: Date) =>
      d instanceof Date ? formatDateForSheet(d) : formatDateForSheet(new Date(d));
    for (let i = 0; i < transactions.length; i++) {
      const txn = transactions[i];
      // Subscriptions: Date | Subscription | USD | USD To Pkr Rate | Amount in Pkr | Total
      const row = [
        txn.date ? fmtDate(txn.date) : '',
        txn.description || txn.beneficiaryName,
        txn.amount,
        null, // USD to PKR rate — to be filled manually or via API
        null, // Amount in PKR
        txn.amount,
      ];
      rows.splice(insertAt + i, 0, row);
      offset++;
    }

    return offset;
  }

  private insertGenericRows(
    rows: unknown[][],
    section: SheetSection,
    transactions: CategorizedTransaction[],
    currentOffset: number,
  ): number {
    let offset = currentOffset;
    const insertAt = section.totalRow + offset;

    const fmtDate = (d: Date) =>
      d instanceof Date ? formatDateForSheet(d) : formatDateForSheet(new Date(d));
    for (let i = 0; i < transactions.length; i++) {
      const txn = transactions[i];
      // Generic: Date | Name/Description | Description | Amount After Tax | Tax | Total
      const row = [
        txn.date ? fmtDate(txn.date) : '',
        txn.canonicalName || txn.beneficiaryName,
        txn.description || null,
        txn.amount,
        txn.taxAmount,
        txn.grossTotal,
      ];
      rows.splice(insertAt + i, 0, row);
      offset++;
    }

    return offset;
  }

  private buildSalaryRow(txn: CategorizedTransaction): unknown[] {
    // Salaries: Date | Employee | Designation | Salary After Tax | Tax | Total | Bonus | CNIC
    // Use date string for proper display in Google Sheets (DD/MM/YYYY)
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

  private preserveColumnWidths(
    source: XLSX.WorkSheet | undefined,
    target: XLSX.WorkSheet,
  ): void {
    if (source?.['!cols']) {
      target['!cols'] = source['!cols'];
    }
    if (source?.['!merges']) {
      target['!merges'] = source['!merges'];
    }
  }

  /**
   * Copy cell formatting (style) from source rows to inserted rows in target.
   * With bottom-to-top insertion, ranges with higher startRow were shifted down
   * by insertions at lower startRow. Compute final positions and copy 's' property.
   */
  private preserveRowFormatting(
    source: XLSX.WorkSheet | undefined,
    target: XLSX.WorkSheet,
    ranges: { startRow: number; count: number; formatSourceRow: number }[],
  ): void {
    if (!source || ranges.length === 0) return;

    // Sort by startRow ascending; ranges with lower startRow shift those with higher
    const sorted = [...ranges].sort((a, b) => a.startRow - b.startRow);
    let shift = 0;
    const maxCol = 20; // Salaries use ~8 cols; copy up to 20 for safety

    for (const r of sorted) {
      const finalStartRow = r.startRow + shift;
      for (let i = 0; i < r.count; i++) {
        const targetRow = finalStartRow + i;
        for (let c = 0; c <= maxCol; c++) {
          const srcAddr = XLSX.utils.encode_cell({ r: r.formatSourceRow, c });
          const tgtAddr = XLSX.utils.encode_cell({ r: targetRow, c });
          const srcCell = source[srcAddr] as { s?: number } | undefined;
          const tgtCell = target[tgtAddr] as { s?: number } | undefined;
          if (srcCell?.s != null && tgtCell) {
            tgtCell.s = srcCell.s;
          }
        }
      }
      shift += r.count;
    }
  }
}
