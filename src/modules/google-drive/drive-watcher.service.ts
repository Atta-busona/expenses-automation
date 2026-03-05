import { Injectable, Logger } from '@nestjs/common';
import { GoogleDriveService, DriveFile } from './google-drive.service';
import { GoogleProcessorService } from '../processor/google-processor.service';
import { extractMonthFromFilename } from '../../common/utils/string.util';

@Injectable()
export class DriveWatcherService {
  private readonly logger = new Logger(DriveWatcherService.name);
  private processedFileIds = new Set<string>();
  private lastCheckTimestamp: string | null = null;
  private watchInterval: ReturnType<typeof setInterval> | null = null;

  constructor(
    private readonly driveService: GoogleDriveService,
    private readonly googleProcessor: GoogleProcessorService,
  ) {}

  /**
   * Start watching a Drive folder for new statement files.
   * When a new file appears, automatically trigger expense processing.
   */
  async startWatching(options: {
    scFolderId: string;
    payoneerFolderId?: string;
    configSpreadsheetId: string;
    expensesSpreadsheetId: string;
    pollIntervalSeconds: number;
    templateSheetName?: string;
    overwrite?: boolean;
  }): Promise<void> {
    this.logger.log(
      `Starting Drive watcher (polling every ${options.pollIntervalSeconds}s)`,
    );
    this.logger.log(`  SC folder: ${options.scFolderId}`);
    if (options.payoneerFolderId) {
      this.logger.log(`  Payoneer folder: ${options.payoneerFolderId}`);
    }

    // Initial scan to mark existing files as "already processed"
    await this.seedProcessedFiles(options.scFolderId);
    if (options.payoneerFolderId) {
      await this.seedProcessedFiles(options.payoneerFolderId);
    }

    this.lastCheckTimestamp = new Date().toISOString();
    this.logger.log(
      `Seeded ${this.processedFileIds.size} existing files. Watching for new uploads...`,
    );

    // Start polling
    this.watchInterval = setInterval(async () => {
      try {
        await this.checkForNewFiles(options);
      } catch (error) {
        this.logger.error(
          `Watch cycle error: ${error instanceof Error ? error.message : error}`,
        );
      }
    }, options.pollIntervalSeconds * 1000);

    // Keep the process alive
    await new Promise(() => {});
  }

  stopWatching(): void {
    if (this.watchInterval) {
      clearInterval(this.watchInterval);
      this.watchInterval = null;
      this.logger.log('Drive watcher stopped');
    }
  }

  private async seedProcessedFiles(folderId: string): Promise<void> {
    const files = await this.driveService.listFilesInFolder(folderId);
    for (const file of files) {
      this.processedFileIds.add(file.id);
    }
  }

  private async checkForNewFiles(options: {
    scFolderId: string;
    payoneerFolderId?: string;
    configSpreadsheetId: string;
    expensesSpreadsheetId: string;
    templateSheetName?: string;
    overwrite?: boolean;
  }): Promise<void> {
    // Check SC folder for new files
    const scFiles = await this.driveService.listFilesInFolder(
      options.scFolderId,
    );
    const newScFiles = scFiles.filter(
      (f) => !this.processedFileIds.has(f.id),
    );

    // Check Payoneer folder for new files
    let newPayoneerFile: DriveFile | undefined;
    if (options.payoneerFolderId) {
      const payFiles = await this.driveService.listFilesInFolder(
        options.payoneerFolderId,
      );
      const newPayFiles = payFiles.filter(
        (f) => !this.processedFileIds.has(f.id),
      );
      if (newPayFiles.length > 0) {
        newPayoneerFile = newPayFiles[0];
      }
    }

    if (newScFiles.length === 0 && !newPayoneerFile) {
      return; // Nothing new
    }

    for (const scFile of newScFiles) {
      const monthFromName = extractMonthFromFilename(scFile.name);
      this.logger.log(
        `New SC statement detected: "${scFile.name}"` +
          (monthFromName ? ` → month: ${monthFromName}` : ''),
      );

      // Match a Payoneer file for the same month
      let matchedPayoneer: DriveFile | undefined;
      if (monthFromName && options.payoneerFolderId) {
        const payFiles = await this.driveService.listFilesInFolder(
          options.payoneerFolderId,
        );
        matchedPayoneer = payFiles.find(
          (f) => extractMonthFromFilename(f.name) === monthFromName,
        );
      } else {
        matchedPayoneer = newPayoneerFile;
      }

      try {
        await this.googleProcessor.processFromDrive({
          configSpreadsheetId: options.configSpreadsheetId,
          expensesSpreadsheetId: options.expensesSpreadsheetId,
          scStatementDriveFile: scFile,
          payoneerStatementDriveFile: matchedPayoneer,
          targetMonth: monthFromName || undefined,
          templateSheetName: options.templateSheetName,
          overwrite: options.overwrite,
        });

        this.processedFileIds.add(scFile.id);
        if (newPayoneerFile) {
          this.processedFileIds.add(newPayoneerFile.id);
        }

        this.logger.log(
          `Successfully processed "${scFile.name}"`,
        );
      } catch (error) {
        this.logger.error(
          `Failed to process "${scFile.name}": ${error instanceof Error ? error.message : error}`,
        );
      }
    }
  }
}
