
export interface SaleRow {
  [key: string]: any;
}

export interface ComparisonDiff {
  uniqueKey: string;
  id: string;
  projectName: string;
  customer: string;
  changes: Array<{
    column: string;
    oldValue: any;
    newValue: any;
  }>;
}

export interface ComparisonResult {
  added: SaleRow[];
  removed: SaleRow[];
  modified: ComparisonDiff[];
  summary?: string;
}

export enum FileType {
  OLD = 'old',
  NEW = 'new'
}
