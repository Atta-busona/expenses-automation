export interface ScBankRawRow {
  'PAYMENT TYPE': string;
  'FILENAME': string | null;
  'PAYMENT REFERENCE': string;
  'FX EXCHANGE RATE': number | null;
  'FILE REFERENCE': string | null;
  'DEBIT DATE': string | number;
  'DEBIT AMOUNT': number;
  'DEBIT AUTHORIZATION CURRENCY EQUIVALENT': number;
  'DEBIT BASE CURRENCY EQUIVALENT': number;
  'BATCH REFERENCE': string;
  'YOUR REFERENCE': string;
  'AUTHORIZED BY': string;
  'AUTHORIZED ON': string;
  'DEBITCCY': string;
  'DEBIT ACCOUNT NUMBER': string;
  'DEBIT ACCOUNT NAME': string;
  'BENEFICIARY NAME': string;
  'BENEFICIARY NICK NAME': string;
  'BENEFICIARY ACCOUNT NUMBER': string;
  'OTHER BANK DETAILS': string | null;
  'PAYMENTCCY': string;
  'AMOUNT': number;
  'PAYMENT DATE': string | number;
  'STATUS': string;
  'NOTES TO SELF': string | null;
}

export interface ScBankTransaction {
  paymentType: string;
  paymentReference: string;
  debitDate: Date;
  amount: number;
  currency: string;
  beneficiaryName: string;
  beneficiaryNickName: string;
  beneficiaryAccount: string;
  paymentDate: Date;
  status: string;
  notesToSelf: string | null;
  authorizedBy: string;
  batchReference: string;
}
