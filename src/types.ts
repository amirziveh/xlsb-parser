export type ProgressCallback = (msg: string, pct: number) => void;

export interface Cell {
  t: 'n' | 's' | 'b' | 'e' | 'blank' | 'f';
  v?: number | string | boolean;
  err?: string;
  ixf?: number;
}

export interface ParsedRow {
  row: number;
  cols: Record<number, Cell>;
}

export interface Sheet {
  name: string;
  rows: ParsedRow[];
  totalCells: number;
}

export interface RawRecord {
  type: string;
  typeNum: number;
  size: number;
  hex: string;
  strings: string[];
}

export interface BinaryDump {
  path: string;
  size: number;
  recCount: number;
  records: RawRecord[];
  typeSummary: Record<string, number>;
}

export interface PivotCacheTable {
  name: string;
  fieldNames: string[];
  rows: (string | number | null)[][];
  rowCount: number;
}

export interface ParsedXlsb {
  sheets: Sheet[];
  sharedStrings: string[];
  xmlFiles: Record<string, string>;
  binaryDumps: BinaryDump[];
  pivotCaches: PivotCacheTable[];
  summary: {
    fileCount: number;
    totalRecords: number;
  };
}
