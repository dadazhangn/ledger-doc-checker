export type ValueKind = 'blank' | 'text' | 'number' | 'boolean' | 'date' | 'formula' | 'error';
export interface CellValue {
  kind: ValueKind;
  raw: string | number | boolean | null;
  text: string;
  formula?: string;
  cached?: CellValue;
}
export interface LedgerCell { address: string; row: number; column: number; value: CellValue; merged?: boolean }
export interface LedgerSheet { name: string; rowCount: number; columnCount: number; cells: LedgerCell[] }
export interface Ledger { name: string; sha256: string; sheets: LedgerSheet[] }
export interface WordParagraph { text: string; location: string; reasons: string[] }
export interface WordCell extends WordParagraph { row: number; column: number; columnSpan: number }
export interface WordTable { index: number; location: string; rows: WordCell[][]; reasons: string[] }
export interface WordDocument {
  name: string;
  sha256: string;
  paragraphs: WordParagraph[];
  tables: WordTable[];
  notices: string[];
}
export interface FileProblem { name: string; message: string }
export interface Batch {
  ledger: Ledger | null;
  documents: WordDocument[];
  errors: FileProblem[];
  files: { name: string; sha256: string }[];
}
export type Comparison = 'exact' | 'trim' | 'date' | 'number';
export interface FieldRule {
  sourceColumn: string;
  targetLabel: string;
  comparison: Comparison;
  tolerance?: number;
  required: boolean;
}
export interface DetailRule {
  sheet: string;
  headerRow: number;
  objectColumn: string;
  keyColumn: string;
  targetKeyHeader: string;
  columns: { sourceColumn: string; targetHeader: string; comparison: Comparison; tolerance?: number; required: boolean }[];
}
export interface Plan {
  version: 1;
  ledger: { sheet: string; headerRow: number; keyColumn: string };
  identityLabel: string;
  bindings: { document: string; key: string }[];
  fields: FieldRule[];
  details: DetailRule | null;
  allowCachedFormula: boolean;
}
export type CheckStatus = 'equal' | 'different' | 'unverifiable';
export interface Evidence { file: string; location: string; value: string }
export interface Finding {
  status: CheckStatus;
  kind: string;
  document: string;
  field: string;
  message: string;
  expected?: Evidence;
  actual?: Evidence;
  notices?: string[];
}
export interface Report {
  schemaVersion: 1;
  engineVersion: string;
  planHash: string;
  files: { name: string; sha256: string }[];
  scope: string;
  plan: Plan;
  findings: Finding[];
  fileErrors: FileProblem[];
  summary: { equal: number; different: number; unverifiable: number; fileErrors: number; documents: number; complete: boolean; verdict: 'equal' | 'different' | 'incomplete' };
}
export interface UploadedFile { name: string; data: string }
