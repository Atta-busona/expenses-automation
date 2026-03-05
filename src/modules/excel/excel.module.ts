import { Module } from '@nestjs/common';
import { ExcelService } from './excel.service';
import { ScBankParser } from './parsers/sc-bank.parser';
import { PayoneerParser } from './parsers/payoneer.parser';

@Module({
  providers: [ExcelService, ScBankParser, PayoneerParser],
  exports: [ExcelService, ScBankParser, PayoneerParser],
})
export class ExcelModule {}
