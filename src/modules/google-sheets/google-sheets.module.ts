import { Module } from '@nestjs/common';
import { GoogleSheetsAdapter } from './google-sheets.adapter';
import { GoogleSheetsExpenseWriter } from './google-sheets-expense.writer';
import { TemplateModule } from '../template/template.module';

@Module({
  imports: [TemplateModule],
  providers: [GoogleSheetsAdapter, GoogleSheetsExpenseWriter],
  exports: [GoogleSheetsAdapter, GoogleSheetsExpenseWriter],
})
export class GoogleSheetsModule {}
