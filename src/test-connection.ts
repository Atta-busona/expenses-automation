import 'reflect-metadata';
import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { GoogleDriveService } from './modules/google-drive/google-drive.service';
import { GoogleSheetsAdapter } from './modules/google-sheets/google-sheets.adapter';

async function test() {
  const app = await NestFactory.createApplicationContext(AppModule, {
    logger: ['log', 'error', 'warn'],
  });

  const drive = app.get(GoogleDriveService);
  const sheets = app.get(GoogleSheetsAdapter);

  console.log('\n=== Testing Google Drive: SC Statements folder ===');
  const scFiles = await drive.listFilesInFolder(
    process.env.SC_STATEMENT_FOLDER_ID!,
  );
  console.log('Files found:', scFiles.length);
  for (const f of scFiles) {
    console.log(`  - ${f.name} (${f.mimeType})`);
  }

  console.log('\n=== Testing Google Drive: Payoneer Statements folder ===');
  const payFiles = await drive.listFilesInFolder(
    process.env.PAYONEER_STATEMENT_FOLDER_ID!,
  );
  console.log('Files found:', payFiles.length);
  for (const f of payFiles) {
    console.log(`  - ${f.name} (${f.mimeType})`);
  }

  console.log('\n=== Testing Google Sheets: Config spreadsheet ===');
  const configSheets = await sheets.getSheetNames(
    process.env.CONFIG_SPREADSHEET_ID!,
  );
  console.log('Sheets:', configSheets);

  console.log('\n=== Testing Google Sheets: Expenses spreadsheet ===');
  const expSheets = await sheets.getSheetNames(
    process.env.EXPENSES_SPREADSHEET_ID!,
  );
  console.log('Sheets:', expSheets);

  console.log('\n=== All connections successful! ===');
  await app.close();
}

// Load .env
import * as fs from 'fs';
import * as path from 'path';
const envPath = path.resolve(__dirname, '../.env');
if (fs.existsSync(envPath)) {
  for (const line of fs.readFileSync(envPath, 'utf-8').split('\n')) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const eqIdx = trimmed.indexOf('=');
    if (eqIdx === -1) continue;
    const key = trimmed.slice(0, eqIdx).trim();
    const value = trimmed.slice(eqIdx + 1).trim();
    if (!process.env[key]) process.env[key] = value;
  }
}

test().catch((e) => {
  console.error('Connection test failed:', e.message);
  process.exit(1);
});
