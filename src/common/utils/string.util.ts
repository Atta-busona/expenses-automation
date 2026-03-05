export function normalizeString(value: unknown): string {
  if (value == null) return '';
  return String(value).trim().toUpperCase();
}

export function normalizeForComparison(value: string): string {
  return value
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .replace(/[^A-Z0-9 ]/g, '');
}

export function extractCardChargeVendor(description: string): string | null {
  const match = description.match(/^Card charge \((.+?)\)/i);
  return match ? match[1].trim() : null;
}

export function isCardCharge(description: string): boolean {
  return /^Card charge \(/i.test(description);
}

export function toBooleanSafe(value: unknown): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    return value.trim().toUpperCase() === 'TRUE';
  }
  return !!value;
}

/**
 * Normalize subheading for comparison: trim and collapse multiple spaces.
 */
export function normalizeSubheading(s: string): string {
  return s.trim().replace(/\s+/g, ' ');
}

/**
 * Extract month label from a filename like "FEB-2026.xlsx" → "FEB-2026".
 * Returns null if the filename doesn't match the expected pattern.
 */
export function extractMonthFromFilename(
  filename: string,
): string | null {
  const match = filename.match(
    /^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)-(\d{4})/i,
  );
  return match ? `${match[1].toUpperCase()}-${match[2]}` : null;
}
