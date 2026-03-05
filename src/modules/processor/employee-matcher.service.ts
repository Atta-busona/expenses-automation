import { Injectable, Logger } from '@nestjs/common';
import { ConfigService } from '../config/config.service';
import { Employee } from '../../common/interfaces/employee.interface';
import { normalizeForComparison } from '../../common/utils/string.util';

@Injectable()
export class EmployeeMatcherService {
  private readonly logger = new Logger(EmployeeMatcherService.name);
  private matchCache = new Map<string, Employee | null>();

  constructor(private readonly configService: ConfigService) {}

  clearCache(): void {
    this.matchCache.clear();
  }

  match(beneficiaryName: string): Employee | null {
    const cacheKey = normalizeForComparison(beneficiaryName);

    if (this.matchCache.has(cacheKey)) {
      return this.matchCache.get(cacheKey) || null;
    }

    // Exact match first
    const exactMatch =
      this.configService.findEmployeeByBeneficiary(beneficiaryName);
    if (exactMatch) {
      this.matchCache.set(cacheKey, exactMatch);
      return exactMatch;
    }

    // Partial token match as fallback: all tokens of employee name
    // must appear in the beneficiary name
    const employees = this.configService.getActiveEmployees();
    const beneficiaryTokens = new Set(cacheKey.split(' '));

    for (const emp of employees) {
      const empTokens = normalizeForComparison(
        emp.beneficiaryNamePrimary,
      ).split(' ');
      const allMatch = empTokens.every((t) => beneficiaryTokens.has(t));
      if (allMatch && empTokens.length >= 2) {
        this.logger.debug(
          `Token match: "${beneficiaryName}" → ${emp.canonicalName}`,
        );
        this.matchCache.set(cacheKey, emp);
        return emp;
      }
    }

    this.logger.warn(`No employee match for: "${beneficiaryName}"`);
    this.matchCache.set(cacheKey, null);
    return null;
  }
}
