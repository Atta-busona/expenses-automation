export interface AppConfig {
  googleServiceAccountKeyPath: string;
  scStatementFolderId: string;
  payoneerStatementFolderId?: string;
  expensesSpreadsheetId: string;
  configSpreadsheetId: string;
  templateSheetName: string;
  watchPollIntervalSeconds: number;
  overwriteExisting: boolean;
}

export function loadAppConfig(): AppConfig {
  return {
    googleServiceAccountKeyPath:
      process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH ||
      './credentials/service-account.json',
    scStatementFolderId: process.env.SC_STATEMENT_FOLDER_ID || '',
    payoneerStatementFolderId: process.env.PAYONEER_STATEMENT_FOLDER_ID,
    expensesSpreadsheetId: process.env.EXPENSES_SPREADSHEET_ID || '',
    configSpreadsheetId: process.env.CONFIG_SPREADSHEET_ID || '',
    templateSheetName: process.env.TEMPLATE_SHEET_NAME || 'Template',
    watchPollIntervalSeconds: parseInt(
      process.env.WATCH_POLL_INTERVAL_SECONDS || '60',
      10,
    ),
    overwriteExisting: process.env.OVERWRITE_EXISTING === 'true',
  };
}
