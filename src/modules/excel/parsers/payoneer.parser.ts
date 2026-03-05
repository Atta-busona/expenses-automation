import { Injectable, Logger } from '@nestjs/common';
import { ExcelService } from '../excel.service';
import {
  PayoneerRawRow,
  PayoneerTransaction,
} from '../../../common/interfaces/payoneer-transaction.interface';
import { parseFlexibleDate } from '../../../common/utils/date.util';
import {
  extractCardChargeVendor,
  isCardCharge,
} from '../../../common/utils/string.util';

@Injectable()
export class PayoneerParser {
  private readonly logger = new Logger(PayoneerParser.name);

  constructor(private readonly excelService: ExcelService) {}

  parse(filePath: string): PayoneerTransaction[] {
    const workbook = this.excelService.readWorkbook(filePath);
    const sheetNames = this.excelService.getSheetNames(workbook);
    const sheetName = sheetNames[0];

    this.logger.log(`Parsing Payoneer statement from sheet: "${sheetName}"`);

    const rawRows = this.excelService.sheetToJson<PayoneerRawRow>(
      workbook,
      sheetName,
    );

    const transactions: PayoneerTransaction[] = [];
    let skipped = 0;

    for (const row of rawRows) {
      const description = (row['Description'] || row['description'] || '')
        .toString()
        .trim();
      const amount = Number(row['Amount'] || row['amount'] || 0);
      const status = (row['Status'] || row['status'] || '')
        .toString()
        .trim()
        .toUpperCase();
      const currency = (row['Currency'] || row['currency'] || 'USD')
        .toString()
        .trim();

      if (status !== 'COMPLETED') {
        skipped++;
        continue;
      }

      // Only process card charges with negative amounts (debits)
      if (!isCardCharge(description) || amount >= 0) {
        skipped++;
        continue;
      }

      const date = parseFlexibleDate(row['Date'] || row['date']);
      const vendorName = extractCardChargeVendor(description);

      transactions.push({
        date: date || new Date(),
        description,
        amount: Math.abs(amount),
        currency,
        status: (row['Status'] || row['status'] || '').toString().trim(),
        vendorName,
        isCardCharge: true,
      });
    }

    this.logger.log(
      `Parsed ${transactions.length} subscription transactions, skipped ${skipped}`,
    );
    return transactions;
  }
}
