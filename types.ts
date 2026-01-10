
export interface SaleRow {
  _raw: any[];
  _colMap: any;
  id: string;
  projectName: string;
  customer: string;
  [key: string]: any;
}

export interface ComparisonDiff {
  uniqueKey: string;
  id: string;
  projectName: string;
  customer: string;
  colBValue?: string; // 儲存 Excel B 欄的資料
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
