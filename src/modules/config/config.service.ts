import { Injectable, Logger } from '@nestjs/common';
import { ExcelService } from '../excel/excel.service';
import { CategoryRule } from '../../common/interfaces/category-rule.interface';
import { Employee } from '../../common/interfaces/employee.interface';
import { ExpenseCategory } from '../../common/enums/category.enum';
import { normalizeSubheading, toBooleanSafe } from '../../common/utils/string.util';
import { parseFlexibleDate } from '../../common/utils/date.util';

@Injectable()
export class ConfigService {
  private readonly logger = new Logger(ConfigService.name);

  private categoryRules: Map<string, CategoryRule> = new Map();
  private employees: Employee[] = [];

  constructor(private readonly excelService: ExcelService) {}

  loadFromWorkbook(configFilePath: string): void {
    const workbook = this.excelService.readWorkbook(configFilePath);
    const categoryRows = this.excelService.sheetToJson<Record<string, unknown>>(
      workbook,
      'CATEGORY_RULES',
    );
    const employeeRows = this.excelService.sheetToJson<Record<string, unknown>>(
      workbook,
      'Employee_Master',
    );
    this.loadFromParsedData(categoryRows, employeeRows);
  }

  /**
   * Load config from pre-parsed row data (used by Google Sheets adapter).
   */
  loadFromParsedData(
    categoryRows: Record<string, unknown>[],
    employeeRows: Record<string, unknown>[],
  ): void {
    this.parseCategoryRules(categoryRows);
    this.parseEmployeeMaster(employeeRows);
  }

  private parseCategoryRules(rows: Record<string, unknown>[]): void {
    this.categoryRules.clear();

    for (const row of rows) {
      const enumKey = (row['enum_key'] || '').toString().trim().toUpperCase();
      if (!enumKey) continue;

      const rule: CategoryRule = {
        enumKey: enumKey as ExpenseCategory,
        targetHeading: (row['target_heading'] || '').toString().trim(),
        isSalary: toBooleanSafe(row['is_salary']),
        requiresEmployeeMatch: toBooleanSafe(row['requires_employee_match']),
        requiresSubheading: toBooleanSafe(row['requires_subheading']),
        allowMultiplePerMonth: toBooleanSafe(row['allow_multiple_per_month']),
        active: toBooleanSafe(row['active']),
        notes: row['notes'] ? row['notes'].toString().trim() : null,
      };

      this.categoryRules.set(enumKey, rule);
    }

    this.logger.log(`Loaded ${this.categoryRules.size} category rules`);
  }

  private parseEmployeeMaster(rows: Record<string, unknown>[]): void {
    this.employees = [];

    for (const row of rows) {
      const empId = (row['employee_id'] || '').toString().trim();
      if (!empId) continue;

      const aliases = (row['beneficiary_name_aliases'] || '')
        .toString()
        .trim();

      const employee: Employee = {
        employeeId: empId,
        canonicalName: (row['canonical_name'] || '').toString().trim(),
        beneficiaryNamePrimary: (row['beneficiary_name_primary'] || '')
          .toString()
          .trim(),
        beneficiaryNameAliases: aliases
          ? aliases.split(',').map((a) => a.trim())
          : [],
        salarySubheading: normalizeSubheading(
          (row['salary_subheading'] || '').toString(),
        ),
        designation: (row['designation'] || '').toString().trim(),
        currency: (row['currency'] || 'PKR').toString().trim(),
        paymentChannel: (row['payment_channel'] || 'bank').toString().trim(),
        salaryNetAfterTax: Number(row['salary_net_after_tax'] || 0),
        taxAmount: Number(row['tax_amount'] || 0),
        salaryGrossTotal: Number(row['salary_gross_total'] || 0),
        defaultBonus: row['default_bonus'] ? Number(row['default_bonus']) : null,
        isActive: toBooleanSafe(row['is_active']),
        validFrom: parseFlexibleDate(row['valid_from']),
        validTo: parseFlexibleDate(row['valid_to']),
        remarks: row['remarks'] ? row['remarks'].toString().trim() : null,
      };

      this.employees.push(employee);
    }

    this.logger.log(
      `Loaded ${this.employees.length} employees from Employee_Master`,
    );
  }

  getCategoryRule(enumKey: string): CategoryRule | undefined {
    return this.categoryRules.get(enumKey.toUpperCase());
  }

  getAllCategoryRules(): Map<string, CategoryRule> {
    return this.categoryRules;
  }

  getActiveEmployees(): Employee[] {
    return this.employees.filter((e) => e.isActive);
  }

  getAllEmployees(): Employee[] {
    return this.employees;
  }

  findEmployeeByBeneficiary(beneficiaryName: string): Employee | undefined {
    const normalized = beneficiaryName.trim().toUpperCase().replace(/\s+/g, ' ');

    return this.employees.find((emp) => {
      if (!emp.isActive) return false;

      const primaryMatch =
        emp.beneficiaryNamePrimary.toUpperCase().replace(/\s+/g, ' ') === normalized;
      if (primaryMatch) return true;

      const canonicalMatch =
        emp.canonicalName.toUpperCase().replace(/\s+/g, ' ') === normalized;
      if (canonicalMatch) return true;

      return emp.beneficiaryNameAliases.some(
        (alias) => alias.toUpperCase().replace(/\s+/g, ' ') === normalized,
      );
    });
  }
}
