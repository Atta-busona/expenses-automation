import 'reflect-metadata';
import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { CliService } from './cli/cli.service';
import { Command } from 'commander';
import * as path from 'path';
import * as fs from 'fs';

async function bootstrap() {
  const app = await NestFactory.createApplicationContext(AppModule, {
    logger: ['log', 'error', 'warn'],
  });

  const program = new Command();

  program
    .name('expense-automation')
    .description(
      'Busona Expense Automation — deterministic expense sheet generation from bank statements',
    )
    .version('1.0.0');

  // ─── LOCAL MODE ──────────────────────────────────────────────
  program
    .command('process')
    .description('Process local Excel files (offline mode)')
    .requiredOption(
      '-c, --config <path>',
      'Path to config workbook (Employee_Master_Template.xlsx)',
    )
    .requiredOption(
      '-s, --sc-statement <path>',
      'Path to SC Bank statement Excel file',
    )
    .option(
      '-p, --payoneer-statement <path>',
      'Path to Payoneer statement Excel file',
    )
    .requiredOption(
      '-e, --expenses-workbook <path>',
      'Path to expenses workbook (BUS-2026-Expenses.xlsx)',
    )
    .requiredOption('-o, --output <path>', 'Output file path')
    .option(
      '-m, --month <label>',
      'Target month (e.g. FEB-2026). Auto-detected if omitted.',
    )
    .option('-t, --template <name>', 'Template sheet name', 'Template')
    .option('--overwrite', 'Overwrite existing monthly sheet', false)
    .action(async (opts) => {
      const cli = app.get(CliService);
      await cli.runLocal({
        configFile: path.resolve(opts.config),
        scStatement: path.resolve(opts.scStatement),
        payoneerStatement: opts.payoneerStatement
          ? path.resolve(opts.payoneerStatement)
          : undefined,
        expensesWorkbook: path.resolve(opts.expensesWorkbook),
        output: path.resolve(opts.output),
        month: opts.month,
        template: opts.template,
        overwrite: opts.overwrite,
      });
    });

  // ─── GOOGLE DRIVE MODE ───────────────────────────────────────
  program
    .command('drive')
    .description(
      'Process the latest statement from Google Drive → Google Sheets',
    )
    .requiredOption(
      '--config-sheet-id <id>',
      'Google Spreadsheet ID for config (Employee_Master_Template)',
    )
    .requiredOption(
      '--expenses-sheet-id <id>',
      'Google Spreadsheet ID for expenses (BUS-2026-Expenses)',
    )
    .requiredOption(
      '--sc-folder-id <id>',
      'Google Drive folder ID for SC Bank statements',
    )
    .option(
      '--payoneer-folder-id <id>',
      'Google Drive folder ID for Payoneer statements',
    )
    .option('-m, --month <label>', 'Target month (auto-detected if omitted)')
    .option('-t, --template <name>', 'Template sheet name', 'Template')
    .option('--overwrite', 'Overwrite existing monthly sheet', false)
    .action(async (opts) => {
      loadEnvFile();
      const cli = app.get(CliService);
      await cli.runFromDrive({
        configSpreadsheetId:
          opts.configSheetId || process.env.CONFIG_SPREADSHEET_ID!,
        expensesSpreadsheetId:
          opts.expensesSheetId || process.env.EXPENSES_SPREADSHEET_ID!,
        scFolderId:
          opts.scFolderId || process.env.SC_STATEMENT_FOLDER_ID!,
        payoneerFolderId:
          opts.payoneerFolderId || process.env.PAYONEER_STATEMENT_FOLDER_ID,
        month: opts.month,
        template: opts.template,
        overwrite: opts.overwrite,
      });
    });

  // ─── WATCH MODE ──────────────────────────────────────────────
  program
    .command('watch')
    .description(
      'Watch a Google Drive folder for new statements and auto-process',
    )
    .option(
      '--config-sheet-id <id>',
      'Google Spreadsheet ID for config',
    )
    .option(
      '--expenses-sheet-id <id>',
      'Google Spreadsheet ID for expenses',
    )
    .option(
      '--sc-folder-id <id>',
      'Google Drive folder ID for SC statements',
    )
    .option(
      '--payoneer-folder-id <id>',
      'Google Drive folder ID for Payoneer statements',
    )
    .option(
      '--poll-interval <seconds>',
      'Poll interval in seconds',
      '60',
    )
    .option('-t, --template <name>', 'Template sheet name', 'Template')
    .option('--overwrite', 'Overwrite existing monthly sheet', false)
    .action(async (opts) => {
      loadEnvFile();
      const cli = app.get(CliService);
      await cli.runWatch({
        configSpreadsheetId:
          opts.configSheetId || process.env.CONFIG_SPREADSHEET_ID!,
        expensesSpreadsheetId:
          opts.expensesSheetId || process.env.EXPENSES_SPREADSHEET_ID!,
        scFolderId:
          opts.scFolderId || process.env.SC_STATEMENT_FOLDER_ID!,
        payoneerFolderId:
          opts.payoneerFolderId || process.env.PAYONEER_STATEMENT_FOLDER_ID,
        pollInterval: parseInt(
          opts.pollInterval ||
            process.env.WATCH_POLL_INTERVAL_SECONDS ||
            '60',
          10,
        ),
        template: opts.template,
        overwrite: opts.overwrite,
      });
    });

  await program.parseAsync(process.argv);
  await app.close();
}

function loadEnvFile(): void {
  const envPath = path.resolve('.env');
  if (!fs.existsSync(envPath)) return;

  const content = fs.readFileSync(envPath, 'utf-8');
  for (const line of content.split('\n')) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eqIdx = trimmed.indexOf('=');
    if (eqIdx === -1) continue;
    const key = trimmed.slice(0, eqIdx).trim();
    const value = trimmed.slice(eqIdx + 1).trim();
    if (!process.env[key]) {
      process.env[key] = value;
    }
  }
}

bootstrap();
