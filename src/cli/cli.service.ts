import { Injectable, Logger } from '@nestjs/common';
import {
  ProcessorService,
  ProcessOptions,
} from '../modules/processor/processor.service';
import { extractMonthFromFilename } from '../common/utils/string.util';
import { GoogleProcessorService } from '../modules/processor/google-processor.service';
import { GoogleDriveService } from '../modules/google-drive/google-drive.service';
import { DriveWatcherService } from '../modules/google-drive/drive-watcher.service';

export interface CliLocalInput {
  configFile: string;
  scStatement: string;
  payoneerStatement?: string;
  expensesWorkbook: string;
  output: string;
  month?: string;
  template?: string;
  overwrite?: boolean;
}

export interface CliDriveInput {
  configSpreadsheetId: string;
  expensesSpreadsheetId: string;
  scFolderId: string;
  payoneerFolderId?: string;
  month?: string;
  template?: string;
  overwrite?: boolean;
}

export interface CliWatchInput {
  configSpreadsheetId: string;
  expensesSpreadsheetId: string;
  scFolderId: string;
  payoneerFolderId?: string;
  pollInterval: number;
  template?: string;
  overwrite?: boolean;
}

@Injectable()
export class CliService {
  private readonly logger = new Logger(CliService.name);

  constructor(
    private readonly processorService: ProcessorService,
    private readonly googleProcessor: GoogleProcessorService,
    private readonly driveService: GoogleDriveService,
    private readonly driveWatcher: DriveWatcherService,
  ) {}

  /**
   * Process local Excel files (original mode).
   */
  async runLocal(input: CliLocalInput): Promise<void> {
    this.logger.log('Expense Automation CLI — Local Mode');

    const options: ProcessOptions = {
      configFilePath: input.configFile,
      scStatementPath: input.scStatement,
      payoneerStatementPath: input.payoneerStatement,
      expensesWorkbookPath: input.expensesWorkbook,
      outputPath: input.output,
      targetMonth: input.month,
      templateSheetName: input.template || 'Template',
      overwrite: input.overwrite ?? false,
    };

    try {
      const result = await this.processorService.process(options);
      this.logger.log(`Done! Output: ${input.output}`);
      this.logger.log(`  Sheet: ${result.monthLabel}`);
      this.logger.log(`  Categorized: ${result.stats.categorizedCount}`);
      this.logger.log(`  Review Required: ${result.stats.reviewCount}`);
    } catch (error) {
      this.logger.error(
        `Processing failed: ${error instanceof Error ? error.message : error}`,
      );
      process.exit(1);
    }
  }

  /**
   * Process the latest statement from Google Drive → Google Sheets.
   */
  async runFromDrive(input: CliDriveInput): Promise<void> {
    this.logger.log('Expense Automation CLI — Google Drive Mode');

    try {
      // Find the latest SC statement in the folder
      const scFiles = await this.driveService.listFilesInFolder(
        input.scFolderId,
      );
      if (scFiles.length === 0) {
        this.logger.error('No statement files found in SC folder');
        process.exit(1);
      }

      const latestSc = scFiles[0]; // Already sorted by modifiedTime desc
      const monthFromName = extractMonthFromFilename(latestSc.name);
      const targetMonth = input.month || monthFromName || undefined;
      this.logger.log(
        `Latest SC statement: "${latestSc.name}"` +
          (monthFromName ? ` → month: ${monthFromName}` : ''),
      );

      // Find matching Payoneer statement (same month if possible)
      let matchedPayoneer = undefined;
      if (input.payoneerFolderId) {
        const payFiles = await this.driveService.listFilesInFolder(
          input.payoneerFolderId,
        );
        if (targetMonth) {
          matchedPayoneer = payFiles.find(
            (f) => extractMonthFromFilename(f.name) === targetMonth,
          );
        }
        if (!matchedPayoneer && payFiles.length > 0) {
          matchedPayoneer = payFiles[0];
        }
        if (matchedPayoneer) {
          this.logger.log(
            `Matched Payoneer statement: "${matchedPayoneer.name}"`,
          );
        }
      }

      const result = await this.googleProcessor.processFromDrive({
        configSpreadsheetId: input.configSpreadsheetId,
        expensesSpreadsheetId: input.expensesSpreadsheetId,
        scStatementDriveFile: latestSc,
        payoneerStatementDriveFile: matchedPayoneer,
        targetMonth,
        templateSheetName: input.template,
        overwrite: input.overwrite,
      });

      this.logger.log(`Done! Sheet "${result.monthLabel}" updated in Google Sheets`);
      this.logger.log(`  Categorized: ${result.stats.categorizedCount}`);
      this.logger.log(`  Review Required: ${result.stats.reviewCount}`);
    } catch (error) {
      this.logger.error(
        `Processing failed: ${error instanceof Error ? error.message : error}`,
      );
      process.exit(1);
    }
  }

  /**
   * Watch a Drive folder and auto-process new statements.
   */
  async runWatch(input: CliWatchInput): Promise<void> {
    this.logger.log('Expense Automation CLI — Watch Mode');
    this.logger.log(
      'Watching for new statement files. Press Ctrl+C to stop.',
    );

    await this.driveWatcher.startWatching({
      scFolderId: input.scFolderId,
      payoneerFolderId: input.payoneerFolderId,
      configSpreadsheetId: input.configSpreadsheetId,
      expensesSpreadsheetId: input.expensesSpreadsheetId,
      pollIntervalSeconds: input.pollInterval,
      templateSheetName: input.template,
      overwrite: input.overwrite,
    });
  }
}
