import { Injectable, Logger } from '@nestjs/common';
import { GoogleAuthService } from '../google-auth/google-auth.service';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

export interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
  createdTime: string;
  modifiedTime: string;
}

@Injectable()
export class GoogleDriveService {
  private readonly logger = new Logger(GoogleDriveService.name);

  constructor(private readonly authService: GoogleAuthService) {}

  /**
   * List files in a Google Drive folder, ordered by most recent first.
   * Only returns spreadsheet and Excel files.
   */
  async listFilesInFolder(folderId: string): Promise<DriveFile[]> {
    const drive = await this.authService.getDriveClient();

    const res = await drive.files.list({
      q: `'${folderId}' in parents and trashed = false and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'application/vnd.ms-excel' or mimeType = 'application/vnd.google-apps.spreadsheet')`,
      fields: 'files(id, name, mimeType, createdTime, modifiedTime)',
      orderBy: 'modifiedTime desc',
      pageSize: 50,
    });

    const files = (res.data.files || []) as DriveFile[];
    this.logger.log(
      `Found ${files.length} files in folder ${folderId}`,
    );
    return files;
  }

  /**
   * List files added after a specific timestamp.
   */
  async listNewFiles(
    folderId: string,
    afterTimestamp: string,
  ): Promise<DriveFile[]> {
    const drive = await this.authService.getDriveClient();

    const res = await drive.files.list({
      q: `'${folderId}' in parents and trashed = false and createdTime > '${afterTimestamp}' and (mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType = 'application/vnd.ms-excel')`,
      fields: 'files(id, name, mimeType, createdTime, modifiedTime)',
      orderBy: 'createdTime desc',
      pageSize: 10,
    });

    return (res.data.files || []) as DriveFile[];
  }

  /**
   * Download a file from Google Drive to a local temp path.
   * Returns the local file path.
   */
  async downloadFile(fileId: string, fileName: string): Promise<string> {
    const drive = await this.authService.getDriveClient();

    const tmpDir = path.join(os.tmpdir(), 'expense-automation');
    if (!fs.existsSync(tmpDir)) {
      fs.mkdirSync(tmpDir, { recursive: true });
    }

    const localPath = path.join(tmpDir, `${fileId}_${fileName}`);

    const res = await drive.files.get(
      { fileId, alt: 'media' },
      { responseType: 'arraybuffer' },
    );

    fs.writeFileSync(localPath, Buffer.from(res.data as ArrayBuffer));
    this.logger.log(`Downloaded "${fileName}" → ${localPath}`);
    return localPath;
  }

  /**
   * Download a Google Sheets file as XLSX to a local temp path.
   */
  async exportSheetAsXlsx(
    fileId: string,
    fileName: string,
  ): Promise<string> {
    const drive = await this.authService.getDriveClient();

    const tmpDir = path.join(os.tmpdir(), 'expense-automation');
    if (!fs.existsSync(tmpDir)) {
      fs.mkdirSync(tmpDir, { recursive: true });
    }

    const localPath = path.join(tmpDir, `${fileId}_${fileName}.xlsx`);

    const res = await drive.files.export(
      {
        fileId,
        mimeType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      { responseType: 'arraybuffer' },
    );

    fs.writeFileSync(localPath, Buffer.from(res.data as ArrayBuffer));
    this.logger.log(`Exported Google Sheet "${fileName}" → ${localPath}`);
    return localPath;
  }

  /**
   * Get a file suitable for local processing — handles both
   * uploaded XLSX and native Google Sheets.
   */
  async getFileForProcessing(file: DriveFile): Promise<string> {
    if (file.mimeType === 'application/vnd.google-apps.spreadsheet') {
      return this.exportSheetAsXlsx(file.id, file.name);
    }
    return this.downloadFile(file.id, file.name);
  }
}
