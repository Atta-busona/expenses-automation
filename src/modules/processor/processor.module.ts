import { Module } from '@nestjs/common';
import { ProcessorService } from './processor.service';
import { GoogleProcessorService } from './google-processor.service';
import { CategorizerService } from './categorizer.service';
import { EmployeeMatcherService } from './employee-matcher.service';
import { ExpenseSheetWriter } from '../excel/writers/expense-sheet.writer';
import { ExcelModule } from '../excel/excel.module';
import { ConfigModule } from '../config/config.module';
import { TemplateModule } from '../template/template.module';
import { GoogleSheetsModule } from '../google-sheets/google-sheets.module';
import { GoogleDriveModule } from '../google-drive/google-drive.module';

@Module({
  imports: [
    ExcelModule,
    ConfigModule,
    TemplateModule,
    GoogleSheetsModule,
    GoogleDriveModule,
  ],
  providers: [
    ProcessorService,
    GoogleProcessorService,
    CategorizerService,
    EmployeeMatcherService,
    ExpenseSheetWriter,
  ],
  exports: [ProcessorService, GoogleProcessorService],
})
export class ProcessorModule {}
