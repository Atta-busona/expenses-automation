import { Injectable, Logger, OnModuleInit } from '@nestjs/common';
import { google } from 'googleapis';
import { JWT } from 'google-auth-library';
import * as fs from 'fs';
import * as path from 'path';

@Injectable()
export class GoogleAuthService implements OnModuleInit {
  private readonly logger = new Logger(GoogleAuthService.name);
  private authClient: JWT | null = null;

  async onModuleInit() {
    // Auth is lazy-loaded on first use — no auto-init required
  }

  async getAuthClient(keyFilePath?: string): Promise<JWT> {
    if (this.authClient) return this.authClient;

    const keyPath = path.resolve(
      keyFilePath ||
        process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH ||
        './credentials/service-account.json',
    );

    if (!fs.existsSync(keyPath)) {
      throw new Error(
        `Google service account key not found at: ${keyPath}\n` +
          'Create a service account in Google Cloud Console, download the JSON key,\n' +
          'and place it at the path specified in GOOGLE_SERVICE_ACCOUNT_KEY_PATH.',
      );
    }

    const keyFile = JSON.parse(fs.readFileSync(keyPath, 'utf-8'));

    this.authClient = new google.auth.JWT({
      email: keyFile.client_email,
      key: keyFile.private_key,
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.readonly',
      ],
    });

    await this.authClient.authorize();
    this.logger.log(
      `Authenticated as service account: ${keyFile.client_email}`,
    );
    return this.authClient;
  }

  async getSheetsClient() {
    const auth = await this.getAuthClient();
    return google.sheets({ version: 'v4', auth });
  }

  async getDriveClient() {
    const auth = await this.getAuthClient();
    return google.drive({ version: 'v3', auth });
  }
}
