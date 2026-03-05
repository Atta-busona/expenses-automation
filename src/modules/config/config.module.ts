import { Module } from '@nestjs/common';
import { ConfigService } from './config.service';
import { ExcelModule } from '../excel/excel.module';

@Module({
  imports: [ExcelModule],
  providers: [ConfigService],
  exports: [ConfigService],
})
export class ConfigModule {}
