export enum ExpenseCategory {
  SALARY = 'SALARY',
  BONUS = 'BONUS',
  ACTIVITY = 'ACTIVITY',
  OPD = 'OPD',
  RENT = 'RENT',
  EQUIPMENT = 'EQUIPMENT',
  VENDOR = 'VENDOR',
  DONATION = 'DONATION',
  SUBSCRIPTION = 'SUBSCRIPTION',
  PROJECT_SHARE = 'PROJECT_SHARE',
  IHSAN_WITHDRAW = 'IHSAN_WITHDRAW',
  TAX_PAYMENT = 'TAX_PAYMENT',
}

export const VALID_CATEGORIES = new Set(Object.values(ExpenseCategory));

export enum ReviewReason {
  MISSING_ENUM = 'Missing enum in Notes to Self',
  INVALID_ENUM = 'Invalid enum — not found in CATEGORY_RULES',
  EMPLOYEE_NOT_FOUND = 'Employee not found for beneficiary',
  DUPLICATE_SALARY = 'Duplicate salary entry for employee this month',
  UNEXPECTED_FORMAT = 'Unexpected row format',
  INACTIVE_CATEGORY = 'Category is inactive in CATEGORY_RULES',
  INVALID_STATUS = 'Transaction status not valid for processing',
  INVALID_SUBHEADING = 'Salary subheading not found in template',
  HEADING_NOT_FOUND = 'Subheading row not found in sheet',
}

export enum TransactionSource {
  SC_BANK = 'SC Bank Statement',
  PAYONEER = 'Payoneer Statement',
}
