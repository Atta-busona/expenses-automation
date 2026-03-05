import { Injectable, Logger } from '@nestjs/common';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

@Injectable()
export class ExcelService {
  private readonly logger = new Logger(ExcelService.name);

  readWorkbook(filePath: string): XLSX.WorkBook {
    const resolved = path.resolve(filePath);
    if (!fs.existsSync(resolved)) {
      throw new Error(`File not found: ${resolved}`);
    }
    this.logger.log(`Reading workbook: ${resolved}`);
    return XLSX.readFile(resolved, { cellDates: false, cellNF: true });
  }

  getSheetNames(workbook: XLSX.WorkBook): string[] {
    return workbook.SheetNames;
  }

  sheetToJson<T = Record<string, unknown>>(
    workbook: XLSX.WorkBook,
    sheetName: string,
  ): T[] {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in workbook`);
    }
    return XLSX.utils.sheet_to_json<T>(sheet, { defval: null });
  }

  sheetToRawRows(
    workbook: XLSX.WorkBook,
    sheetName: string,
  ): unknown[][] {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in workbook`);
    }
    return XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: null,
    });
  }

  getSheet(workbook: XLSX.WorkBook, sheetName: string): XLSX.WorkSheet {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found in workbook`);
    }
    return sheet;
  }

  duplicateSheet(
    workbook: XLSX.WorkBook,
    sourceSheetName: string,
    targetSheetName: string,
  ): void {
    const sourceSheet = this.getSheet(workbook, sourceSheetName);
    const cloned = JSON.parse(JSON.stringify(sourceSheet));

    if (workbook.SheetNames.includes(targetSheetName)) {
      this.logger.warn(`Sheet "${targetSheetName}" already exists, will be replaced`);
      delete workbook.Sheets[targetSheetName];
      const idx = workbook.SheetNames.indexOf(targetSheetName);
      workbook.SheetNames.splice(idx, 1);
    }

    workbook.SheetNames.push(targetSheetName);
    workbook.Sheets[targetSheetName] = cloned;
    this.logger.log(`Duplicated "${sourceSheetName}" → "${targetSheetName}"`);
  }

  writeWorkbook(workbook: XLSX.WorkBook, filePath: string): void {
    const resolved = path.resolve(filePath);
    const dir = path.dirname(resolved);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    XLSX.writeFile(workbook, resolved);
    this.logger.log(`Written workbook: ${resolved}`);
  }

  createSheet(data: unknown[][]): XLSX.WorkSheet {
    return XLSX.utils.aoa_to_sheet(data);
  }

  addSheetToWorkbook(
    workbook: XLSX.WorkBook,
    sheetName: string,
    sheet: XLSX.WorkSheet,
  ): void {
    if (workbook.SheetNames.includes(sheetName)) {
      delete workbook.Sheets[sheetName];
      const idx = workbook.SheetNames.indexOf(sheetName);
      workbook.SheetNames.splice(idx, 1);
    }
    workbook.SheetNames.push(sheetName);
    workbook.Sheets[sheetName] = sheet;
  }
}
