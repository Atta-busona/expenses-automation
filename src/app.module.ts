import { Module } from '@nestjs/common';
import { ExcelModule } from './modules/excel/excel.module';
import { ConfigModule } from './modules/config/config.module';
import { ProcessorModule } from './modules/processor/processor.module';
import { TemplateModule } from './modules/template/template.module';
import { GoogleAuthModule } from './modules/google-auth/google-auth.module';
import { GoogleSheetsModule } from './modules/google-sheets/google-sheets.module';
import { GoogleDriveModule } from './modules/google-drive/google-drive.module';
import { CliModule } from './cli/cli.module';

@Module({
  imports: [
    GoogleAuthModule,
    ExcelModule,
    ConfigModule,
    ProcessorModule,
    TemplateModule,
    GoogleSheetsModule,
    GoogleDriveModule,
    CliModule,
  ],
})
export class AppModule {}
