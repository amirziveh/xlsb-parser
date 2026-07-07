// Streaming handle for memory-efficient iteration of large XLSB sheets.
//
// `openXlsb(data)` unzips the workbook, parses workbook/shared-strings/styles
// (the cheap, shared parts), and yields back a handle. The handle exposes
// `iterSheetRows(index)` which returns an async generator that yields
// complete ParsedRow objects one at a time — directly from the records()
// iterator with one row's worth of cell buffering at a time, so memory stays
// O(cells_per_row) instead of O(total_rows).

import { unzipSync } from 'fflate';
import type { ParseOptions, ProgressCallback, ParsedRow, Sheet, StylesTable, PivotCacheCell, PivotCacheTable, PivotCacheSummary } from './types.js';
import { XlsbSizeError } from './types.js';
import {
  records,
  readU16,
  readU32,
  BRT_ROW_HEADER,
  BRT_CELL_BLANK,
  BRT_FMLA_ERROR,
  BRT_SHORT_BLANK,
  BRT_SHORT_ISST,
  BRT_BEGIN_PIVOT_CACHE_RECORDS,
} from './record-stream.js';
import { parseWorkbook } from './workbook.js';
import { parseSharedStrings } from './shared-strings.js';
import { parseStyles } from './styles.js';
import { readCell, readShortCell, applyDateMeta } from './sheet.js';
import { parseDefinition, parsePivotCache, streamPivotRows } from './pivot-cache.js';

export interface XlsbHandle {
  sheetNames: string[];
  sharedStrings: string[];
  styles: StylesTable | null;
  pivotCaches: PivotCacheSummary[];
  /** Iterate rows of `sheetIndex`-th sheet, one ParsedRow at a time. */
  iterSheetRows(sheetIndex: number, options?: IterOptions): AsyncGenerator<ParsedRow>;
  /** Convenience: drain the entire sheet (still streams internally). */
  collectSheet(sheetIndex: number, options?: IterOptions): Promise<Sheet>;
  /** Stream rows from a pivot cache by numeric index or name. */
  iterPivotCacheRows(
    indexOrName: number | string,
    options?: IterOptions,
  ): AsyncGenerator<PivotCacheCell[]>;
  /** Drain a pivot cache into an in-memory PivotCacheTable. */
  collectPivotCache(indexOrName: number | string): Promise<PivotCacheTable>;
}

export interface IterOptions {
  maxRows?: number;
  onProgress?: ProgressCallback;
}

function tick(): Promise<void> {
  return new Promise((r) => setTimeout(r, 0));
}

function normalizeOptions(arg: ParseOptions | ProgressCallback | undefined): ParseOptions {
  if (arg === undefined) return {};
  if (typeof arg === 'function') return { onProgress: arg };
  return arg;
}

export async function openXlsb(
  data: ArrayBuffer | Uint8Array,
  options?: ParseOptions | ProgressCallback,
): Promise<XlsbHandle> {
  const opts = normalizeOptions(options);
  const onProgress = opts.onProgress;
  const maxZipBytes = opts.maxZipBytes;

  onProgress?.('Decompressing ZIP...', 0);
  await new Promise((r) => setTimeout(r, 50));
  const u8 = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
  let zip: Record<string, Uint8Array>;
  try {
    zip = unzipSync(u8);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new Error('ZIP decompression failed: ' + msg, { cause: e });
  }

  if (maxZipBytes !== undefined) {
    let total = 0;
    for (const k of Object.keys(zip)) total += zip[k].length;
    if (total > maxZipBytes) {
      throw new XlsbSizeError(
        `Decompressed ZIP size ${total} bytes exceeds maxZipBytes limit ${maxZipBytes}`,
        maxZipBytes,
        total,
      );
    }
  }

  const wb = zip['xl/workbook.bin'];
  if (!wb) throw new Error('xl/workbook.bin not found');

  onProgress?.('Parsing workbook...', 5);
  const sheetNames = parseWorkbook(wb);
  await tick();

  let sharedStrings: string[] = [];
  if (zip['xl/sharedStrings.bin']) {
    onProgress?.('Parsing shared strings...', 10);
    sharedStrings = parseSharedStrings(zip['xl/sharedStrings.bin']);
    await tick();
  }

  let styles: StylesTable | null = null;
  if (zip['xl/styles.bin']) {
    onProgress?.('Parsing styles...', 12);
    try {
      styles = parseStyles(zip['xl/styles.bin']);
    } catch {
      styles = null;
    }
    await tick();
  }

  // Eagerly parse pivot-cache definition headers if opted in.
  const pivotDefs: { name: string; def: Uint8Array; recs: Uint8Array }[] = [];
  if (opts.parsePivotCaches) {
    const defPaths = Object.keys(zip)
      .filter(k => /^xl\/pivotCache\/pivotCacheDefinition\d+\.bin$/.test(k))
      .sort((a, b) => {
        const na = parseInt(a.match(/(\d+)\.bin$/)![1], 10);
        const nb = parseInt(b.match(/(\d+)\.bin$/)![1], 10);
        return na - nb;
      });
    for (const defPath of defPaths) {
      const num = defPath.match(/(\d+)\.bin$/)![1];
      const def = zip[defPath];
      const recs = zip[`xl/pivotCache/pivotCacheRecords${num}.bin`];
      if (def && recs) {
        pivotDefs.push({ name: `PivotCache${num}`, def, recs });
      }
    }
  }

  // Capture per-sheet bytes by reference so the iterator can walk them lazily.
  const sheetBytes: (Uint8Array | null)[] = sheetNames.map((_, i) => {
    return zip[`xl/worksheets/sheet${i + 1}.bin`] ?? null;
  });

  const handle: XlsbHandle = {
    sheetNames,
    sharedStrings,
    styles,
    pivotCaches: pivotDefs.map(d => {
      try {
        const defs = parseDefinition(d.def);
        const pcs: PivotCacheSummary = {
          name: d.name,
          fieldNames: defs.map(b => b.name),
          fields: defs.map(b => ({ name: b.name, isSrc: b.isSrc, kind: b.kind, sharedItems: b.sharedItems })),
          rowCount: 0,
        };
        for (const r of records(d.recs)) {
          if (r.type === BRT_BEGIN_PIVOT_CACHE_RECORDS && r.data.length >= 4) {
            pcs.rowCount = readU32(r.data, 0);
            break;
          }
        }
        return pcs;
      } catch {
        return { name: d.name, fieldNames: [], fields: [], rowCount: 0 };
      }
    }),
    async *iterSheetRows(
      sheetIndex: number,
      iterOpts: IterOptions = {},
    ): AsyncGenerator<ParsedRow> {
      const bytes = sheetBytes[sheetIndex];
      if (!bytes) return;
      const maxRows = iterOpts.maxRows;
      const onProgressIter = iterOpts.onProgress;
      const ss = this.sharedStrings;
      const styles = this.styles;

      let curRow: ParsedRow | null = null;
      let prevCol = -1;
      let yielded = 0;
      let lastProgressPct = -1;

      for (const r of records(bytes)) {
        const d = r.data;
        if (r.type === BRT_ROW_HEADER && d.length >= 4) {
          // Yield the previous row (now complete) and start a new one.
          if (curRow) {
            yield curRow;
            yielded++;
            if (onProgressIter && yielded % 1000 === 0) {
              const pct = Math.min(99, Math.floor((yielded / 100000) * 100));
              if (pct !== lastProgressPct) {
                onProgressIter(`Row ${yielded}`, pct);
                lastProgressPct = pct;
              }
              await tick();
            }
            if (maxRows !== undefined && yielded >= maxRows) return;
          }
          curRow = { row: readU32(d, 0), cols: {} };
          prevCol = -1;
          continue;
        }
        if (!curRow) continue;

        // Cell records attach to the current row, exactly like parseSheet.
        if (r.type >= BRT_CELL_BLANK && r.type <= BRT_FMLA_ERROR) {
          const col = readU32(d, 0);
          const ixf =
            d.length >= 6 ? d[4] | (d[5] << 8) | (d[5] & 0x80 ? 0xffff0000 : 0) : undefined;
          const cell = readCell(r.type, d, 8, ss);
          if (cell) {
            cell.ixf = ixf;
            applyDateMeta(cell, ixf, styles);
            curRow.cols[col] = cell;
          }
          prevCol = col;
        } else if (r.type >= BRT_SHORT_BLANK && r.type <= BRT_SHORT_ISST) {
          const col = prevCol + 1;
          const ixf = d.length >= 4 ? readU16(d, 2) : undefined;
          const cell = readShortCell(r.type, d, 4, ss);
          if (cell) {
            cell.ixf = ixf;
            applyDateMeta(cell, ixf, styles);
            curRow.cols[col] = cell;
          }
          prevCol = col;
        }
      }
      // Yield the last row (in case the stream ends without another row header).
      if (curRow) {
        yield curRow;
        yielded++;
      }
      if (onProgressIter) onProgressIter(`Done (${yielded} rows)`, 100);
    },
    async collectSheet(sheetIndex: number, iterOpts: IterOptions = {}): Promise<Sheet> {
      const name = sheetNames[sheetIndex] ?? `Sheet${sheetIndex + 1}`;
      const rows: ParsedRow[] = [];
      for await (const row of this.iterSheetRows(sheetIndex, iterOpts)) rows.push(row);
      const totalCells = rows.reduce((a, r) => a + Object.keys(r.cols).length, 0);
      return { name, rows, totalCells };
    },
    async *iterPivotCacheRows(
      indexOrName: number | string,
      iterOpts: IterOptions = {},
    ): AsyncGenerator<PivotCacheCell[]> {
      const idx = typeof indexOrName === 'number'
        ? indexOrName
        : pivotDefs.findIndex(d => d.name === indexOrName);
      const d = pivotDefs[idx];
      if (!d) return;
      const maxRows = iterOpts.maxRows;
      let yielded = 0;
      for (const row of streamPivotRows(d.def, d.recs)) {
        yield row;
        yielded++;
        if (maxRows !== undefined && yielded >= maxRows) return;
      }
    },
    async collectPivotCache(indexOrName: number | string): Promise<PivotCacheTable> {
      const idx = typeof indexOrName === 'number'
        ? indexOrName
        : pivotDefs.findIndex(d => d.name === indexOrName);
      const d = pivotDefs[idx];
      if (!d) throw new Error(`Pivot cache not found: ${indexOrName}`);
      return parsePivotCache(d.name, d.def, d.recs);
    },
  };

  return handle;
}
