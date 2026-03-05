import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '../config/config.service';
import { ScBankTransaction } from '../../common/interfaces/sc-transaction.interface';
import { PayoneerTransaction } from '../../common/interfaces/payoneer-transaction.interface';
import {
  CategorizedTransaction,
} from '../../common/interfaces/processing-result.interface';
import { ReviewItem } from '../../common/interfaces/review-item.interface';
import {
  ExpenseCategory,
  ReviewReason,
  TransactionSource,
  VALID_CATEGORIES,
} from '../../common/enums/category.enum';

@Injectable()
export class CategorizerService {
  private readonly logger = new Logger(CategorizerService.name);

  constructor(private readonly configService: ConfigService) {}

  categorizeScTransaction(
    txn: ScBankTransaction,
  ): { categorized: CategorizedTransaction | null; review: ReviewItem | null } {
    const rawEnum = txn.notesToSelf
      ? txn.notesToSelf.trim().toUpperCase()
      : null;

    if (!rawEnum) {
      return {
        categorized: null,
        review: this.buildReview(txn, ReviewReason.MISSING_ENUM),
      };
    }

    if (!VALID_CATEGORIES.has(rawEnum as ExpenseCategory)) {
      return {
        categorized: null,
        review: this.buildReview(txn, ReviewReason.INVALID_ENUM),
      };
    }

    const rule = this.configService.getCategoryRule(rawEnum);
    if (!rule) {
      return {
        categorized: null,
        review: this.buildReview(txn, ReviewReason.INVALID_ENUM),
      };
    }

    if (!rule.active) {
      return {
        categorized: null,
        review: this.buildReview(txn, ReviewReason.INACTIVE_CATEGORY),
      };
    }

    let employee = null;
    if (rule.requiresEmployeeMatch) {
      employee = this.configService.findEmployeeByBeneficiary(
        txn.beneficiaryName,
      );
      if (!employee) {
        return {
          categorized: null,
          review: this.buildReview(txn, ReviewReason.EMPLOYEE_NOT_FOUND),
        };
      }
    }

    // For salary, use Employee_Master figures; for others, use bank amount
    const useMasterSalary = rule.isSalary && employee;
    const amount = useMasterSalary ? employee!.salaryNetAfterTax : txn.amount;
    const taxAmount = useMasterSalary ? employee!.taxAmount : 0;
    const grossTotal = useMasterSalary ? employee!.salaryGrossTotal : txn.amount;

    const categorized: CategorizedTransaction = {
      date: txn.paymentDate,
      beneficiaryName: txn.beneficiaryName,
      amount,
      currency: txn.currency,
      category: rawEnum,
      targetHeading: rule.targetHeading,
      subheading: employee?.salarySubheading || null,
      designation: employee?.designation || null,
      employeeId: employee?.employeeId || null,
      canonicalName: employee?.canonicalName || null,
      taxAmount,
      grossTotal,
      bonus: 0,
      cnic: null,
      description: null,
      source: TransactionSource.SC_BANK,
    };

    return { categorized, review: null };
  }

  categorizePayoneerTransaction(
    txn: PayoneerTransaction,
  ): CategorizedTransaction {
    return {
      date: txn.date,
      beneficiaryName: txn.vendorName || txn.description,
      amount: txn.amount,
      currency: txn.currency,
      category: ExpenseCategory.SUBSCRIPTION,
      targetHeading: 'Subscriptions',
      subheading: null,
      designation: null,
      employeeId: null,
      canonicalName: null,
      taxAmount: 0,
      grossTotal: txn.amount,
      bonus: 0,
      cnic: null,
      description: txn.vendorName || txn.description,
      source: TransactionSource.PAYONEER,
    };
  }

  private buildReview(
    txn: ScBankTransaction,
    reason: ReviewReason,
  ): ReviewItem {
    return {
      date: txn.paymentDate,
      beneficiary: txn.beneficiaryName,
      amount: txn.amount,
      currency: txn.currency,
      notes: txn.notesToSelf,
      reason,
      sourceFile: TransactionSource.SC_BANK,
    };
  }
}
