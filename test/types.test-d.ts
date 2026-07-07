// Type-level tests for the public API. Run via `npm run test:types`.
// These are compile-time assertions only — no runtime code.

import {
  parseXlsb,
  openXlsb,
  XlsbSizeError,
  type ParsedXlsb,
  type Cell,
  type ParsedRow,
  type Sheet,
  type ParseOptions,
  type ProgressCallback,
  type StylesTable,
  type XlsbHandle,
  type IterOptions,
} from '../src/index.js';

export const _compileTimeChecks: void = (() => {
  // parseXlsb accepts ArrayBuffer | Uint8Array and returns Promise<ParsedXlsb>.
  const _p1: Promise<ParsedXlsb> = parseXlsb(new ArrayBuffer(0));
  const _p2: Promise<ParsedXlsb> = parseXlsb(new Uint8Array(0));

  // New options bag form.
  const _p3: Promise<ParsedXlsb> = parseXlsb(new ArrayBuffer(0), {
    onProgress: (_msg, _pct) => {},
    maxZipBytes: 1_000_000,
    maxRowsPerSheet: 10_000,
    dumpBinaries: true,
    readXml: true,
    parsePivotCaches: true,
  });

  // openXlsb returns a handle.
  const _h: Promise<XlsbHandle> = openXlsb(new ArrayBuffer(0));

  // Handle exposes sheetNames + sharedStrings + styles + iterator.
  const handle = {} as XlsbHandle;
  const _names: string[] = handle.sheetNames;
  const _ss: string[] = handle.sharedStrings;
  const _st: StylesTable | null = handle.styles;
  const _iter: AsyncGenerator<ParsedRow> = handle.iterSheetRows(0);
  const _iter2: AsyncGenerator<ParsedRow> = handle.iterSheetRows(0, { maxRows: 10 });
  const _collect: Promise<Sheet> = handle.collectSheet(0);

  // Cell shape: the new date metadata fields are optional.
  const cell = {} as Cell;
  const _numFmtId: number | undefined = cell.numFmtId;
  const _isDate: boolean | undefined = cell.isDate;
  const _dateValue: string | undefined = cell.dateValue;
  const _ixf: number | undefined = cell.ixf;

  // ParseOptions fields.
  const opts = {} as ParseOptions;
  const _cb: ProgressCallback | undefined = opts.onProgress;
  const _mzb: number | undefined = opts.maxZipBytes;

  // XlsbSizeError is a real class with limit + actual fields.
  const err = new XlsbSizeError('test', 100, 200);
  const _limit: number = err.limit;
  const _actual: number = err.actual;
})();
