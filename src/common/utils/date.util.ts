/**
 * Excel serial date epoch: Jan 0, 1900 (with the Lotus 1-2-3 leap year bug).
 * Serial 1 = Jan 1, 1900.
 */
const EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30));

export function excelSerialToDate(serial: number): Date {
  const ms = serial * 86400000;
  return new Date(EXCEL_EPOCH.getTime() + ms);
}

export function isExcelSerial(value: unknown): value is number {
  return typeof value === 'number' && value > 40000 && value < 60000;
}

export function parseFlexibleDate(value: unknown): Date | null {
  if (value == null) return null;

  if (isExcelSerial(value)) {
    return excelSerialToDate(value);
  }

  if (value instanceof Date) return value;

  if (typeof value === 'string') {
    // DD/MM/YYYY format (SC Bank uses this)
    const ddmmyyyy = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (ddmmyyyy) {
      const [, day, month, year] = ddmmyyyy;
      return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    }

    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
  }

  return null;
}

export function formatMonthLabel(date: Date): string {
  const months = [
    'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
    'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC',
  ];
  return `${months[date.getMonth()]}-${date.getFullYear()}`;
}

export function formatDateForSheet(date: Date): string {
  const day = date.getDate().toString().padStart(2, '0');
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

export function dateToExcelSerial(date: Date): number {
  const utcDate = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return (utcDate - EXCEL_EPOCH.getTime()) / 86400000;
}
