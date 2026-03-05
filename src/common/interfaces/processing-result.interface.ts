import { ReviewItem } from './review-item.interface';

export interface CategorizedTransaction {
  date: Date;
  beneficiaryName: string;
  amount: number;
  currency: string;
  category: string;
  targetHeading: string;
  subheading: string | null;
  designation: string | null;
  employeeId: string | null;
  canonicalName: string | null;
  taxAmount: number;
  grossTotal: number;
  bonus: number;
  cnic: string | null;
  description: string | null;
  source: string;
}

export interface ProcessingResult {
  monthLabel: string;
  categorized: CategorizedTransaction[];
  reviewItems: ReviewItem[];
  stats: ProcessingStats;
}

export interface ProcessingStats {
  totalScTransactions: number;
  totalPayoneerTransactions: number;
  categorizedCount: number;
  reviewCount: number;
  skippedCount: number;
  byCategory: Record<string, number>;
}
