export type ProgressCallback = (msg: string, pct: number) => void;

// Options bag (P4). Replaces the legacy `parseXlsb(data, onProgress?)` form.
// For backwards compatibility through 1.x, a function passed as the 2nd arg
// is treated as `onProgress`. Will be removed at 2.0.
export interface ParseOptions {
  onProgress?: ProgressCallback;
  /** Throw XlsbSizeError if decompressed ZIP total exceeds this many bytes. */
  maxZipBytes?: number;
  /** Stop parsing a sheet after this many rows (rest of sheet skipped). */
  maxRowsPerSheet?: number;
  /** If true, populate `binaryDumps` (debug record dumper). Default: false. */
  dumpBinaries?: boolean;
  /** If true, populate `xmlFiles` with raw XML/rels content. Default: false. */
  readXml?: boolean;
  /** If true, parse xl/pivotCache/*.bin. Default: false (opt-in). */
  parsePivotCaches?: boolean;
}

export interface Cell {
  t: 'n' | 's' | 'b' | 'e' | 'blank' | 'f';
  v?: number | string | boolean;
  err?: string;
  ixf?: number;
  /** Numeric format ID resolved from the styles table (P4). */
  numFmtId?: number;
  /** True when numFmtId is a date/time format (P4). */
  isDate?: boolean;
  /** ISO 8601 string when isDate && t === 'n' (P4); v stays the raw serial. */
  dateValue?: string;
}

export type PivotCacheCell =
  | { t: 's'; v: string }
  | { t: 'n'; v: number }
  | { t: 'd'; v: string; serial?: number }
  | { t: 'b'; v: boolean }
  | { t: 'e'; v: string }
  | { t: 'blank' };

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

export type PivotCacheFieldKind =
  | 'indexed' | 'number' | 'date' | 'string' | 'boolean' | 'error' | 'blank';

export interface PivotCacheField {
  name: string;
  isSrc: boolean;
  kind: PivotCacheFieldKind;
  sharedItems: (PivotCacheCell | null)[];
}

export interface PivotCacheTable {
  name: string;
  fieldNames: string[];
  fields: PivotCacheField[];
  rows: PivotCacheCell[][];
  rowCount: number;
  recordCount: number;
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

export interface PivotCacheSummary {
  name: string;
  fieldNames: string[];
  fields: PivotCacheField[];
  rowCount: number;
}

// Raised when maxZipBytes or maxRowsPerSheet caps are exceeded. Distinct
// from parse errors so consumers can differentiate "too big" from "broken".
export class XlsbSizeError extends Error {
  override name = 'XlsbSizeError';
  readonly limit: number;
  readonly actual: number;
  constructor(message: string, limit: number, actual: number) {
    super(message);
    this.limit = limit;
    this.actual = actual;
  }
}

// Style table — populated from xl/styles.bin. Maps cell iStyleRef → numFmtId.
// Used to mark cells as dates (when numFmtId is a date format).
export interface StylesTable {
  // cellXfs[i] = numFmtId applied to cells whose ixf === i.
  cellXfs: number[];
  // Custom numFmt strings registered via numFmts (built-ins resolved by ID).
  numFmts: Map<number, string>;
}
