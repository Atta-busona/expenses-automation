import { ExpenseCategory } from '../enums/category.enum';

export interface CategoryRule {
  enumKey: ExpenseCategory;
  targetHeading: string;
  isSalary: boolean;
  requiresEmployeeMatch: boolean;
  requiresSubheading: boolean;
  allowMultiplePerMonth: boolean;
  active: boolean;
  notes: string | null;
}
