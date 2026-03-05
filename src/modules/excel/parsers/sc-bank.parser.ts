import { Injectable, Logger } from '@nestjs/common';
import { ExcelService } from '../excel.service';
import {
  ScBankRawRow,
  ScBankTransaction,
} from '../../../common/interfaces/sc-transaction.interface';
import { parseFlexibleDate } from '../../../common/utils/date.util';

const VALID_STATUSES = new Set([
  'PROCESSED BY BANK',
  'CREDIT SUCCESSFUL',
]);

@Injectable()
export class ScBankParser {
  private readonly logger = new Logger(ScBankParser.name);

  constructor(private readonly excelService: ExcelService) {}

  parse(filePath: string): ScBankTransaction[] {
    const workbook = this.excelService.readWorkbook(filePath);
    const sheetNames = this.excelService.getSheetNames(workbook);
    const sheetName = sheetNames[0];

    this.logger.log(`Parsing SC Bank statement from sheet: "${sheetName}"`);

    const rawRows = this.excelService.sheetToJson<ScBankRawRow>(
      workbook,
      sheetName,
    );

    const transactions: ScBankTransaction[] = [];
    let skipped = 0;

    for (const row of rawRows) {
      const status = (row['STATUS'] || '').toString().trim().toUpperCase();
      if (!VALID_STATUSES.has(status)) {
        skipped++;
        continue;
      }

      const amount = Number(row['AMOUNT'] || row['DEBIT AMOUNT'] || 0);
      if (amount <= 0) {
        skipped++;
        continue;
      }

      const debitDate = parseFlexibleDate(row['DEBIT DATE']);
      const paymentDate = parseFlexibleDate(row['PAYMENT DATE']);

      if (!debitDate) {
        this.logger.warn(
          `Skipping row with unparseable date: ${row['DEBIT DATE']}`,
        );
        skipped++;
        continue;
      }

      transactions.push({
        paymentType: (row['PAYMENT TYPE'] || '').toString().trim(),
        paymentReference: (row['PAYMENT REFERENCE'] || '').toString().trim(),
        debitDate,
        amount,
        currency: (row['DEBITCCY'] || row['PAYMENTCCY'] || 'PKR')
          .toString()
          .trim(),
        beneficiaryName: (row['BENEFICIARY NAME'] || '').toString().trim(),
        beneficiaryNickName: (row['BENEFICIARY NICK NAME'] || '')
          .toString()
          .trim(),
        beneficiaryAccount: (row['BENEFICIARY ACCOUNT NUMBER'] || '')
          .toString()
          .trim(),
        paymentDate: paymentDate || debitDate,
        status: (row['STATUS'] || '').toString().trim(),
        notesToSelf: row['NOTES TO SELF']
          ? row['NOTES TO SELF'].toString().trim()
          : null,
        authorizedBy: (row['AUTHORIZED BY'] || '').toString().trim(),
        batchReference: (row['BATCH REFERENCE'] || '').toString().trim(),
      });
    }

    this.logger.log(
      `Parsed ${transactions.length} valid transactions, skipped ${skipped}`,
    );
    return transactions;
  }
}
