import { ReviewReason, TransactionSource } from '../enums/category.enum';

export interface ReviewItem {
  date: Date | null;
  beneficiary: string;
  amount: number;
  currency: string;
  notes: string | null;
  reason: ReviewReason;
  sourceFile: TransactionSource;
}
