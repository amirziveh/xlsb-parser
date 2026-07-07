import type { Cell, ParsedRow, StylesTable } from './types.js';
import { XlsbSizeError } from './types.js';
import {
  records, readU16, readU32, readF64, readWideString, readRichString, decodeRk,
  BRT_ROW_HEADER, BRT_CELL_BLANK, BRT_CELL_RK, BRT_CELL_ERROR, BRT_CELL_BOOL,
  BRT_CELL_REAL, BRT_CELL_ST, BRT_CELL_ISST,
  BRT_FMLA_STRING, BRT_FMLA_NUM, BRT_FMLA_BOOL, BRT_FMLA_ERROR,
  BRT_SHORT_BLANK, BRT_SHORT_RK, BRT_SHORT_ERROR, BRT_SHORT_BOOL,
  BRT_SHORT_REAL, BRT_SHORT_ST, BRT_SHORT_ISST,
  ERRORS,
} from './record-stream.js';
import { isDateFormatId, isDateNumFmtString, numFmtIdToSerialConvert } from './styles.js';

export interface ParseSheetOptions {
  /** Stop after parsing this many rows (rest of sheet skipped). */
  maxRows?: number;
  /** Styles table — when present, cells get numFmtId/isDate/dateValue. */
  styles?: StylesTable | null;
}

// Parse xl/worksheets/sheetN.bin → ParsedRow[].
// Two record encodings exist per cell: a long form (BRT_CELL_*) that carries
// its own column index, and a short form (BRT_SHORT_*) that uses prevCol+1.
export function parseSheet(
  data: Uint8Array,
  ss: string[],
  opts: ParseSheetOptions = {},
): ParsedRow[] {
  const rows: ParsedRow[] = [];
  const styles = opts.styles ?? null;
  const maxRows = opts.maxRows;
  let curRow: ParsedRow | null = null;
  let prevCol = -1;

  for (const r of records(data)) {
    const d = r.data;
    if (r.type === BRT_ROW_HEADER && d.length >= 4) {
      if (maxRows !== undefined && rows.length >= maxRows) {
        throw new XlsbSizeError(
          `Sheet exceeded maxRows=${maxRows} limit`,
          maxRows, rows.length,
        );
      }
      curRow = { row: readU32(d, 0), cols: {} };
      rows.push(curRow);
      prevCol = -1;
      continue;
    }
    if (!curRow) continue;

    if (r.type >= BRT_CELL_BLANK && r.type <= BRT_FMLA_ERROR) {
      const col = readU32(d, 0);
      // iStyleRef is a signed 16-bit value at d[4..5]; sign bit lives in d[5].
      const ixf = d.length >= 6 ? (d[4] | (d[5] << 8)) | (d[5] & 0x80 ? 0xFFFF0000 : 0) : undefined;
      const cell = readCell(r.type, d, 8, ss);
      if (cell) { cell.ixf = ixf; applyDateMeta(cell, ixf, styles); curRow.cols[col] = cell; }
      prevCol = col;
    } else if (r.type >= BRT_SHORT_BLANK && r.type <= BRT_SHORT_ISST) {
      const col = prevCol + 1;
      const ixf = d.length >= 4 ? readU16(d, 2) : undefined;
      const cell = readShortCell(r.type, d, 4, ss);
      if (cell) { cell.ixf = ixf; applyDateMeta(cell, ixf, styles); curRow.cols[col] = cell; }
      prevCol = col;
    }
  }
  return rows;
}

// Resolve numFmtId from the styles table. If the format is a date/time format,
// populate isDate=true and dateValue (ISO 8601) for numeric cells using the
// Excel 1900 epoch (1899-12-30 + serial*86400000ms).
export function applyDateMeta(cell: Cell, ixf: number | undefined, styles: StylesTable | null): void {
  if (!styles || ixf === undefined || ixf < 0) return;
  // ixf indexes into cellXfs; cellXfs[ixf] is the numFmtId applied to the cell.
  // Old signed iStyleRef can be negative; clamp.
  const idx = ixf & 0xFFFF;
  if (idx >= styles.cellXfs.length) return;
  const numFmtId = styles.cellXfs[idx];
  cell.numFmtId = numFmtId;

  const customFmt = styles.numFmts.get(numFmtId);
  const isDate =
    (customFmt !== undefined && isDateNumFmtString(customFmt)) ||
    isDateFormatId(numFmtId);
  if (!isDate) return;
  cell.isDate = true;
  if (cell.t === 'n' && typeof cell.v === 'number') {
    cell.dateValue = numFmtIdToSerialConvert(cell.v);
  }
}

// Long-form cell reader. `off` points at the value payload (after the 8-byte
// col/iStyleRef/reserved header). Every case validates d.length before
// reading; truncated records (the result of partial writes or fuzzing) are
// dropped rather than reading past r.data's end.
export function readCell(type: number, d: Uint8Array, off: number, ss: string[]): Cell | null {
  switch (type) {
    case BRT_CELL_BLANK: return { t: 'blank' };
    case BRT_CELL_RK:
      if (off + 4 > d.length) return null;
      return { t: 'n', v: decodeRk(readU32(d, off)) };
    case BRT_CELL_REAL:
      if (off + 8 > d.length) return null;
      return { t: 'n', v: readF64(d, off) };
    case BRT_CELL_ISST:
      if (off + 4 > d.length) return null;
      return { t: 's', v: ss[readU32(d, off)] ?? `[SST#${readU32(d, off)}]` };
    case BRT_CELL_BOOL:
      if (off + 1 > d.length) return null;
      return { t: 'b', v: d[off] !== 0 };
    case BRT_CELL_ERROR:
      if (off + 1 > d.length) return null;
      return { t: 'e', err: ERRORS[d[off]] ?? `#ERR(${d[off]})` };
    case BRT_CELL_ST:
      if (off + 5 > d.length) return null;
      return { t: 's', v: readRichString(d, off) };
    case BRT_FMLA_NUM:
      if (off + 8 > d.length) return null;
      return { t: 'n', v: readF64(d, off) };
    case BRT_FMLA_STRING:
      if (off + 4 > d.length) return null;
      return { t: 's', v: readWideString(d, off) };
    case BRT_FMLA_BOOL:
      if (off + 1 > d.length) return null;
      return { t: 'b', v: d[off] !== 0 };
    case BRT_FMLA_ERROR:
      if (off + 1 > d.length) return null;
      return { t: 'e', err: ERRORS[d[off]] ?? `#ERR(${d[off]})` };
    default: return null;
  }
}

// Short-form reader. Same length checks, narrower value types (no formula
// variants exist in short form).
export function readShortCell(type: number, d: Uint8Array, off: number, ss: string[]): Cell | null {
  switch (type) {
    case BRT_SHORT_BLANK: return { t: 'blank' };
    case BRT_SHORT_RK:
      if (off + 4 > d.length) return null;
      return { t: 'n', v: decodeRk(readU32(d, off)) };
    case BRT_SHORT_ERROR:
      if (off + 1 > d.length) return null;
      return { t: 'e', err: ERRORS[d[off]] ?? `#ERR(${d[off]})` };
    case BRT_SHORT_BOOL:
      if (off + 1 > d.length) return null;
      return { t: 'b', v: d[off] !== 0 };
    case BRT_SHORT_REAL:
      if (off + 8 > d.length) return null;
      return { t: 'n', v: readF64(d, off) };
    case BRT_SHORT_ST:
      if (off + 5 > d.length) return null;
      return { t: 's', v: readRichString(d, off) };
    case BRT_SHORT_ISST:
      if (off + 4 > d.length) return null;
      return { t: 's', v: ss[readU32(d, off)] ?? `[SST#${readU32(d, off)}]` };
    default: return null;
  }
}
