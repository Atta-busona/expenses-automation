import { Module } from '@nestjs/common';
import { TemplateService } from './template.service';
import { ExcelModule } from '../excel/excel.module';

@Module({
  imports: [ExcelModule],
  providers: [TemplateService],
  exports: [TemplateService],
})
export class TemplateModule {}
