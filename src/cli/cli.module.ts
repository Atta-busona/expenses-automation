import { Module } from '@nestjs/common';
import { CliService } from './cli.service';
import { ProcessorModule } from '../modules/processor/processor.module';
import { GoogleDriveModule } from '../modules/google-drive/google-drive.module';
import { DriveWatcherService } from '../modules/google-drive/drive-watcher.service';

@Module({
  imports: [ProcessorModule, GoogleDriveModule],
  providers: [CliService, DriveWatcherService],
  exports: [CliService],
})
export class CliModule {}
