export interface PayoneerRawRow {
  Date: string | number;
  Description: string;
  Amount: number;
  Currency: string;
  Status: string;
  [key: string]: unknown;
}

export interface PayoneerTransaction {
  date: Date;
  description: string;
  amount: number;
  currency: string;
  status: string;
  vendorName: string | null;
  isCardCharge: boolean;
}
