import { Injectable, Logger } from '@nestjs/common';
import * as XLSX from 'xlsx';
import { ExcelService } from '../excel/excel.service';
import { normalizeSubheading } from '../../common/utils/string.util';

export interface SheetSection {
  headingRow: number;
  headingText: string;
  headerRow: number;
  headers: string[];
  dataStartRow: number;
  totalRow: number;
  subheadings: SubheadingRange[];
}

export interface SubheadingRange {
  label: string;
  row: number;
  /** Column index where subheading text lives (0=Date, 1=Employee) */
  col: number;
  dataStartRow: number;
  dataEndRow: number;
}

@Injectable()
export class TemplateService {
  private readonly logger = new Logger(TemplateService.name);

  constructor(private readonly excelService: ExcelService) {}

  ensureMonthlySheet(
    workbook: XLSX.WorkBook,
    monthLabel: string,
    templateSheetName: string = 'Template',
    overwrite: boolean = false,
  ): string {
    // Fallback: if "Template" not found, use first sheet (common when sheet has different name)
    let actualTemplate = templateSheetName;
    if (!workbook.SheetNames.includes(templateSheetName) && workbook.SheetNames.length > 0) {
      actualTemplate = workbook.SheetNames[0];
      this.logger.warn(
        `Sheet "${templateSheetName}" not found, using first sheet "${actualTemplate}" as template`,
      );
    }

    const exists = workbook.SheetNames.includes(monthLabel);

    if (exists && overwrite) {
      this.logger.log(`Overwriting existing sheet: "${monthLabel}"`);
      delete workbook.Sheets[monthLabel];
      const idx = workbook.SheetNames.indexOf(monthLabel);
      workbook.SheetNames.splice(idx, 1);
    } else if (exists && !overwrite) {
      let suffix = 2;
      let altName = `${monthLabel}-${suffix}`;
      while (workbook.SheetNames.includes(altName)) {
        suffix++;
        altName = `${monthLabel}-${suffix}`;
      }
      this.logger.log(
        `Sheet "${monthLabel}" exists, creating "${altName}" instead`,
      );
      this.excelService.duplicateSheet(workbook, actualTemplate, altName);
      return altName;
    }

    this.excelService.duplicateSheet(workbook, actualTemplate, monthLabel);
    return monthLabel;
  }

  parseSheetSections(
    workbook: XLSX.WorkBook,
    sheetName: string,
  ): SheetSection[] {
    const rawRows = this.excelService.sheetToRawRows(workbook, sheetName);
    return this.parseSheetSectionsFromRows(rawRows, sheetName);
  }

  /**
   * Parse sections from raw row data (for use with Google Sheets API when
   * we read directly from the sheet without XLSX).
   */
  parseSheetSectionsFromRows(
    rawRows: unknown[][],
    sheetName: string = 'Sheet',
  ): SheetSection[] {
    const sections: SheetSection[] = [];

    const knownHeadings = new Set([
      'SALARIES',
      'TEAM ACTIVITIES & BONUSES',
      'TEAM OPD',
      "IHSAN'S WITHDRAWAL",
      'IHSAN WITHDRAWL',
      'PROJECT SHARES & COMMISSIONS',
      'SUBSCRIPTIONS',
      'DK RENT',
      'VENDORS',
      'EQUIPMENTS',
      'EQUIPMENT',
      'DONATIONS',
      'TAX PAYMENTS',
      'GRAND TOTAL',
    ]);

    // Known salary sub-headings (department names)
    const salarySubheadings = new Set([
      'CEO',
      'UI/UX DESIGNERS',
      'PROJECT MANAGERS',
      'WEBFLOW DEVELOPERS',
      'FRONTEND & BACKEND ENGINEERS',
      'MARKETING',
      'HR',
      'SALES TEAM',
      'TAXATION',
    ]);

    for (let i = 0; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (!row || !row[0]) continue;

      const cellValue = String(row[0]).trim().toUpperCase();

      // Check if this is a known section heading
      if (!knownHeadings.has(cellValue)) continue;
      if (cellValue === 'GRAND TOTAL') continue;

      const headingRow = i;
      const headingText = String(row[0]).trim();

      // Next row is typically the column header
      const headerRow = i + 1;
      const headers = rawRows[headerRow]
        ? rawRows[headerRow]
            .map((c) => (c != null ? String(c).trim() : ''))
            .filter(Boolean)
        : [];

      // Find the Total row for this section (scan full row for merged cells)
      let totalRow = -1;
      let nextSectionRow = -1;
      for (let j = headerRow + 1; j < rawRows.length; j++) {
        const r = rawRows[j];
        if (this.cellInRowEquals(r, 'TOTAL')) {
          totalRow = j;
          break;
        }
        if (this.rowContainsKnownHeading(r, knownHeadings)) {
          nextSectionRow = j;
          totalRow = j - 1;
          break;
        }
      }

      // For sections without a Total row (e.g. Tax Payments at end),
      // find the last non-empty data row and insert after it.
      if (totalRow === -1 || (nextSectionRow === -1 && totalRow >= rawRows.length - 2)) {
        let lastDataRow = headerRow;
        const limit = Math.min(headerRow + 200, rawRows.length);
        for (let j = headerRow + 1; j < limit; j++) {
          const r = rawRows[j];
          if (!r) continue;
          const hasData = r.some(
            (c) => c != null && String(c).trim() !== '',
          );
          if (hasData) lastDataRow = j;
          else if (j - lastDataRow > 3) break; // stop after 3 consecutive empty rows
        }
        totalRow = lastDataRow + 1;
      }

      // Parse sub-headings for salary section.
      // Merged cells: heading text may appear in any column of the merged range.
      // Scan the ENTIRE row for each cell; if any cell matches a valid subheading, mark that row.
      const subheadings: SubheadingRange[] = [];
      if (cellValue === 'SALARIES') {
        for (let j = headerRow + 1; j < totalRow; j++) {
          const r = rawRows[j];
          if (!r) continue;
          const { rawLabel, col } = this.findSubheadingInRow(r, salarySubheadings);
          if (!rawLabel) continue;

          const dataStart = j + 1;
          let dataEnd = totalRow - 1;
          for (let k = j + 1; k < totalRow; k++) {
            const sr = rawRows[k];
            if (!sr) continue;
            const hit = this.findSubheadingOrTotalInRow(sr, salarySubheadings);
            if (hit) {
              dataEnd = k - 1;
              break;
            }
          }

          subheadings.push({
            label: rawLabel,
            row: j,
            col,
            dataStartRow: dataStart,
            dataEndRow: dataEnd,
          });
        }
      }

      if (cellValue === 'SALARIES' && subheadings.length === 0) {
        this.logger.warn(
          'Salaries section has no subheadings (UI/UX Designers, Project Managers, etc.). ' +
            'Ensure the Template sheet has these rows with department names.',
        );
      } else if (cellValue === 'SALARIES' && subheadings.length > 0) {
        this.logger.log(
          `Found ${subheadings.length} salary subheadings: ${subheadings.map((s) => s.label).join(', ')}`,
        );
      }

      sections.push({
        headingRow,
        headingText,
        headerRow,
        headers: headers as string[],
        dataStartRow: headerRow + 1,
        totalRow,
        subheadings,
      });
    }

    this.logger.log(
      `Parsed ${sections.length} sections from "${sheetName}": ${sections.map((s) => s.headingText).join(', ')}`,
    );
    return sections;
  }

  /**
   * Get the set of valid salary subheadings (exact text from template, normalized).
   * Used to strictly validate Employee_Master salary_subheading before insertion.
   */
  getValidSalarySubheadings(
    sections: SheetSection[],
  ): Set<string> {
    const salarySection = sections.find((s) =>
      s.headingText.toLowerCase().includes('salar'),
    );
    const valid = new Set<string>();
    if (!salarySection?.subheadings?.length) return valid;
    for (const sub of salarySection.subheadings) {
      valid.add(normalizeSubheading(sub.label));
    }
    // CEO is always valid (inserted before Total)
    valid.add(normalizeSubheading('CEO'));
    return valid;
  }

  findSectionByHeading(
    sections: SheetSection[],
    targetHeading: string,
  ): SheetSection | undefined {
    const normalizedTarget = targetHeading.trim().toUpperCase();

    return sections.find((s) => {
      const normalizedSection = s.headingText.trim().toUpperCase();
      return (
        normalizedSection === normalizedTarget ||
        normalizedSection.includes(normalizedTarget) ||
        normalizedTarget.includes(normalizedSection)
      );
    });
  }

  /**
   * Find subheading by exact normalized match only (no fuzzy).
   * Both labels are normalized with trim + collapse spaces for strict matching.
   */
  findSubheading(
    section: SheetSection,
    subheadingLabel: string,
  ): SubheadingRange | undefined {
    const normalized = normalizeSubheading(subheadingLabel);

    return section.subheadings.find((sh) => {
      const sheetNorm = normalizeSubheading(sh.label.replace(/\s*\(\d+\)\s*$/, ''));
      return sheetNorm === normalized;
    });
  }

  /** Check if any cell in row equals value (case-insensitive). */
  private cellInRowEquals(row: unknown[] | undefined, value: string): boolean {
    const v = value.trim().toUpperCase();
    const maxCol = Math.min((row?.length ?? 0) + 10, 20);
    for (let c = 0; c < maxCol; c++) {
      const cell = row?.[c] != null ? String(row[c]).trim().toUpperCase() : '';
      if (cell === v) return true;
    }
    return false;
  }

  /** Check if any cell in row is a known section heading. */
  private rowContainsKnownHeading(
    row: unknown[] | undefined,
    known: Set<string>,
  ): boolean {
    const maxCol = Math.min((row?.length ?? 0) + 10, 20);
    for (let c = 0; c < maxCol; c++) {
      const cell = row?.[c] != null ? String(row[c]).trim().toUpperCase() : '';
      if (cell && (known.has(cell) || cell === 'GRAND TOTAL')) return true;
    }
    return false;
  }

  /**
   * Scan entire row for a salary subheading. Handles merged cells where text
   * may appear in any column. Returns { rawLabel, col } or { rawLabel: null }.
   */
  private findSubheadingInRow(
    row: unknown[],
    known: Set<string>,
  ): { rawLabel: string | null; col: number } {
    const maxCol = Math.min((row?.length ?? 0) + 10, 20);
    for (let c = 0; c < maxCol; c++) {
      const cellVal = row?.[c] != null ? String(row[c]).trim() : '';
      if (!cellVal) continue;
      const stripped = cellVal.replace(/\s*\(\d+\)\s*$/, '').trim().toUpperCase();
      if (this.matchesSalarySubheading(stripped, known)) {
        return { rawLabel: cellVal, col: c };
      }
    }
    return { rawLabel: null, col: 0 };
  }

  /**
   * Scan entire row for a salary subheading OR "TOTAL". Returns true if found.
   */
  private findSubheadingOrTotalInRow(
    row: unknown[],
    known: Set<string>,
  ): boolean {
    const maxCol = Math.min((row?.length ?? 0) + 10, 20);
    for (let c = 0; c < maxCol; c++) {
      const cellVal = row?.[c] != null ? String(row[c]).trim() : '';
      if (!cellVal) continue;
      const stripped = cellVal.replace(/\s*\(\d+\)\s*$/, '').trim().toUpperCase();
      if (stripped === 'TOTAL' || this.matchesSalarySubheading(stripped, known)) {
        return true;
      }
    }
    return false;
  }

  private matchesSalarySubheading(
    stripped: string,
    known: Set<string>,
  ): boolean {
    if (known.has(stripped)) return true;
    // Fuzzy: check if any known subheading is a close match
    for (const k of known) {
      if (this.fuzzyMatch(stripped, k)) return true;
    }
    return false;
  }

  /**
   * Simple token-overlap fuzzy match: at least 80% of tokens must match.
   * Handles typos like "Backedn" vs "Backend" by checking token similarity.
   */
  private fuzzyMatch(a: string, b: string): boolean {
    const tokensA = a.split(/[\s&]+/).filter(Boolean);
    const tokensB = b.split(/[\s&]+/).filter(Boolean);
    if (tokensA.length === 0 || tokensB.length === 0) return false;

    let matches = 0;
    for (const ta of tokensA) {
      for (const tb of tokensB) {
        if (ta === tb || this.levenshtein(ta, tb) <= 2) {
          matches++;
          break;
        }
      }
    }

    const ratio = matches / Math.max(tokensA.length, tokensB.length);
    return ratio >= 0.75;
  }

  private levenshtein(a: string, b: string): number {
    const m = a.length;
    const n = b.length;
    const dp: number[][] = Array.from({ length: m + 1 }, () =>
      Array(n + 1).fill(0),
    );
    for (let i = 0; i <= m; i++) dp[i][0] = i;
    for (let j = 0; j <= n; j++) dp[0][j] = j;
    for (let i = 1; i <= m; i++) {
      for (let j = 1; j <= n; j++) {
        dp[i][j] = Math.min(
          dp[i - 1][j] + 1,
          dp[i][j - 1] + 1,
          dp[i - 1][j - 1] + (a[i - 1] !== b[j - 1] ? 1 : 0),
        );
      }
    }
    return dp[m][n];
  }
}
