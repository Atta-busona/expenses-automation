export interface Employee {
  employeeId: string;
  canonicalName: string;
  beneficiaryNamePrimary: string;
  beneficiaryNameAliases: string[];
  salarySubheading: string;
  designation: string;
  currency: string;
  paymentChannel: string;
  salaryNetAfterTax: number;
  taxAmount: number;
  salaryGrossTotal: number;
  defaultBonus: number | null;
  isActive: boolean;
  validFrom: Date | null;
  validTo: Date | null;
  remarks: string | null;
}
