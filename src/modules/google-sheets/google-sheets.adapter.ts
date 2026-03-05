import { Injectable, Logger } from '@nestjs/common';
import { GoogleAuthService } from '../google-auth/google-auth.service';
import { sheets_v4 } from 'googleapis';

@Injectable()
export class GoogleSheetsAdapter {
  private readonly logger = new Logger(GoogleSheetsAdapter.name);

  constructor(private readonly authService: GoogleAuthService) {}

  private async getClient(): Promise<sheets_v4.Sheets> {
    return this.authService.getSheetsClient();
  }

  async getSheetNames(spreadsheetId: string): Promise<string[]> {
    const sheets = await this.getClient();
    const res = await sheets.spreadsheets.get({ spreadsheetId });
    return (res.data.sheets || []).map((s) => s.properties?.title || '');
  }

  async getSheetIdByName(
    spreadsheetId: string,
    sheetName: string,
  ): Promise<number | null> {
    const sheets = await this.getClient();
    const res = await sheets.spreadsheets.get({ spreadsheetId });
    const sheet = (res.data.sheets || []).find(
      (s) => s.properties?.title === sheetName,
    );
    return sheet?.properties?.sheetId ?? null;
  }

  /**
   * Read all data from a sheet as a 2D array.
   */
  async readSheet(
    spreadsheetId: string,
    sheetName: string,
  ): Promise<unknown[][]> {
    const sheets = await this.getClient();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: sheetName,
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'SERIAL_NUMBER',
    });

    return (res.data.values || []) as unknown[][];
  }

  /**
   * Read a sheet as key-value objects using the first row as headers.
   */
  async readSheetAsJson(
    spreadsheetId: string,
    sheetName: string,
  ): Promise<Record<string, unknown>[]> {
    const rows = await this.readSheet(spreadsheetId, sheetName);
    if (rows.length < 2) return [];

    const headers = rows[0].map((h) => String(h || '').trim());
    return rows.slice(1).map((row) => {
      const obj: Record<string, unknown> = {};
      headers.forEach((h, i) => {
        obj[h] = (row as unknown[])[i] ?? null;
      });
      return obj;
    });
  }

  /**
   * Write data to a sheet, replacing existing content.
   * Clear removes values only; formatting (colors, etc.) is preserved.
   */
  async writeToSheet(
    spreadsheetId: string,
    sheetName: string,
    data: unknown[][],
  ): Promise<void> {
    const sheets = await this.getClient();

    // Clear values (formatting preserved per Sheets API)
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: sheetName,
    });

    // Write new data
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: data },
    });

    this.logger.log(
      `Wrote ${data.length} rows to "${sheetName}" in spreadsheet ${spreadsheetId}`,
    );
  }

  /**
   * Update a specific range without clearing the sheet.
   */
  async updateRange(
    spreadsheetId: string,
    sheetName: string,
    range: string,
    data: unknown[][],
  ): Promise<void> {
    const sheets = await this.getClient();
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!${range}`,
      valueInputOption: 'RAW',
      requestBody: { values: data },
    });
  }

  /**
   * Append rows to the end of a sheet's data.
   */
  async appendToSheet(
    spreadsheetId: string,
    sheetName: string,
    data: unknown[][],
  ): Promise<void> {
    const sheets = await this.getClient();

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetName}!A1`,
      valueInputOption: 'RAW',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: data },
    });

    this.logger.log(
      `Appended ${data.length} rows to "${sheetName}"`,
    );
  }

  /**
   * Insert rows at a specific position in a sheet.
   */
  async insertRowsAt(
    spreadsheetId: string,
    sheetName: string,
    startRow: number,
    data: unknown[][],
  ): Promise<void> {
    const sheets = await this.getClient();
    const sheetId = await this.getSheetIdByName(spreadsheetId, sheetName);
    if (sheetId === null) {
      throw new Error(`Sheet "${sheetName}" not found`);
    }

    // Insert empty rows first
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            insertDimension: {
              range: {
                sheetId,
                dimension: 'ROWS',
                startIndex: startRow,
                endIndex: startRow + data.length,
              },
              inheritFromBefore: true,
            },
          },
        ],
      },
    });

    // Write data into the new rows
    const range = `${sheetName}!A${startRow + 1}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'RAW',
      requestBody: { values: data },
    });
  }

  /**
   * Duplicate a sheet within the same spreadsheet.
   */
  async duplicateSheet(
    spreadsheetId: string,
    sourceSheetName: string,
    newTitle: string,
  ): Promise<number> {
    const sheets = await this.getClient();
    const sourceSheetId = await this.getSheetIdByName(
      spreadsheetId,
      sourceSheetName,
    );

    if (sourceSheetId === null) {
      throw new Error(
        `Source sheet "${sourceSheetName}" not found in spreadsheet`,
      );
    }

    // Check if target already exists
    const existingId = await this.getSheetIdByName(spreadsheetId, newTitle);
    if (existingId !== null) {
      this.logger.warn(`Sheet "${newTitle}" already exists, deleting first`);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [{ deleteSheet: { sheetId: existingId } }],
        },
      });
    }

    const res = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            duplicateSheet: {
              sourceSheetId,
              newSheetName: newTitle,
            },
          },
        ],
      },
    });

    const newSheetId =
      res.data.replies?.[0]?.duplicateSheet?.properties?.sheetId ?? null;

    if (newSheetId !== null) {
      await this.moveSheetToEnd(spreadsheetId, newSheetId);
    }

    this.logger.log(
      `Duplicated "${sourceSheetName}" → "${newTitle}" (sheetId: ${newSheetId})`,
    );
    return newSheetId ?? 0;
  }

  /**
   * Move a sheet to the end (rightmost position).
   */
  async moveSheetToEnd(
    spreadsheetId: string,
    sheetId: number,
  ): Promise<void> {
    const sheets = await this.getClient();
    const res = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetCount = (res.data.sheets || []).length;
    const lastIndex = Math.max(0, sheetCount - 1);

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: { sheetId, index: lastIndex },
              fields: 'index',
            },
          },
        ],
      },
    });
    this.logger.log(`Moved sheet ${sheetId} to end (index ${lastIndex})`);
  }

  /**
   * Delete a sheet by name.
   */
  async deleteSheet(
    spreadsheetId: string,
    sheetName: string,
  ): Promise<void> {
    const sheetId = await this.getSheetIdByName(spreadsheetId, sheetName);
    if (sheetId === null) return;

    const sheets = await this.getClient();
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: [{ deleteSheet: { sheetId } }],
      },
    });
    this.logger.log(`Deleted sheet "${sheetName}"`);
  }

  /**
   * Check if a sheet exists in the spreadsheet.
   */
  async sheetExists(
    spreadsheetId: string,
    sheetName: string,
  ): Promise<boolean> {
    const names = await this.getSheetNames(spreadsheetId);
    return names.includes(sheetName);
  }
}
