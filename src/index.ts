import { unzipSync } from 'fflate';
import type {
  ProgressCallback, ParsedXlsb, ParseOptions,
} from './types.js';
import { XlsbSizeError, type StylesTable } from './types.js';
import {
  dec8,
} from './record-stream.js';
import { parseWorkbook } from './workbook.js';
import { parseSharedStrings } from './shared-strings.js';
import { parseSheet } from './sheet.js';
import { dumpBinary } from './dump.js';
import { parsePivotCache } from './pivot-cache.js';
import { parseStyles, isDateFormatId, isDateNumFmtString } from './styles.js';

// Re-export the public types so consumers can `import type { ... } from 'xlsb-parser'`.
export type {
  ProgressCallback, Cell, ParsedRow, Sheet, RawRecord, BinaryDump,
  PivotCacheTable, ParsedXlsb, ParseOptions, StylesTable,
} from './types.js';
export { XlsbSizeError } from './types.js';
export { openXlsb } from './handle.js';
export type { XlsbHandle, IterOptions } from './handle.js';

function tick(): Promise<void> {
  return new Promise(r => setTimeout(r, 0));
}

// Normalise the legacy 2nd-arg-as-function form into ParseOptions. Behaviour:
//   parseXlsb(data)                       → {}
//   parseXlsb(data, fn)                  → { onProgress: fn }  (deprecated, 1.x only)
//   parseXlsb(data, opts)                → opts
function normalizeOptions(
  arg: ParseOptions | ProgressCallback | undefined,
): ParseOptions {
  if (arg === undefined) return {};
  if (typeof arg === 'function') return { onProgress: arg };
  return arg;
}

export async function parseXlsb(
  data: ArrayBuffer | Uint8Array,
  options?: ParseOptions | ProgressCallback,
): Promise<ParsedXlsb> {
  const opts = normalizeOptions(options);
  const onProgress = opts.onProgress;
  const dumpBinaries = opts.dumpBinaries === true;
  const readXml = opts.readXml === true;
  const parsePivotCaches = opts.parsePivotCaches === true;
  const maxRowsPerSheet = opts.maxRowsPerSheet;
  const maxZipBytes = opts.maxZipBytes;

  onProgress?.('Decompressing ZIP...', 0);
  // Yield so browser paints the log before synchronous decompression
  await new Promise(r => setTimeout(r, 50));
  const u8 = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
  let zip: Record<string, Uint8Array>;
  try {
    zip = unzipSync(u8);
  } catch (e: any) {
    throw new Error('ZIP decompression failed: ' + (e?.message || e));
  }

  if (maxZipBytes !== undefined) {
    let total = 0;
    for (const k of Object.keys(zip)) total += zip[k].length;
    if (total > maxZipBytes) {
      throw new XlsbSizeError(
        `Decompressed ZIP size ${total} bytes exceeds maxZipBytes limit ${maxZipBytes}`,
        maxZipBytes, total,
      );
    }
  }

  const out: ParsedXlsb = {
    sheets: [], sharedStrings: [], xmlFiles: {}, binaryDumps: [], pivotCaches: [],
    summary: { fileCount: 0, totalRecords: 0 },
  };

  const wb = zip['xl/workbook.bin'];
  if (!wb) throw new Error('xl/workbook.bin not found');

  onProgress?.('Parsing workbook...', 5);
  const sheetNames = parseWorkbook(wb);
  await tick();

  if (zip['xl/sharedStrings.bin']) {
    onProgress?.('Parsing shared strings...', 10);
    out.sharedStrings = parseSharedStrings(zip['xl/sharedStrings.bin']);
    await tick();
  }

  // Styles — parsed opportunistically; absent styles.bin → null.
  let styles: StylesTable | null = null;
  if (zip['xl/styles.bin']) {
    onProgress?.('Parsing styles...', 12);
    try { styles = parseStyles(zip['xl/styles.bin']); } catch { styles = null; }
    await tick();
  }

  for (let i = 0; i < sheetNames.length; i++) {
    const key = `xl/worksheets/sheet${i + 1}.bin`;
    const sd = zip[key];
    if (sd) {
      onProgress?.(`Sheet "${sheetNames[i]}"...`, 15 + Math.round((i / sheetNames.length) * 20));
      const rows = parseSheet(sd, out.sharedStrings, { maxRows: maxRowsPerSheet, styles });
      const totalCells = rows.reduce((a, r) => a + Object.keys(r.cols).length, 0);
      out.sheets.push({ name: sheetNames[i], rows, totalCells });
      await tick();
    }
  }

  // Pivot caches — opt-in since P4 (default: skip).
  if (parsePivotCaches) {
    const pcd1 = zip['xl/pivotCache/pivotCacheDefinition1.bin'];
    const pcd2 = zip['xl/pivotCache/pivotCacheDefinition2.bin'];
    const pcr1 = zip['xl/pivotCache/pivotCacheRecords1.bin'];
    const pcr2 = zip['xl/pivotCache/pivotCacheRecords2.bin'];
    if (pcd1 && pcr1) {
      onProgress?.('Pivot cache 1...', 33);
      out.pivotCaches.push(parsePivotCache('PivotCache1', pcd1, pcr1));
      await tick();
    }
    if (pcd2 && pcr2) {
      onProgress?.('Pivot cache 2...', 34);
      out.pivotCaches.push(parsePivotCache('PivotCache2', pcd2, pcr2));
      await tick();
    }
  }

  // Binary dumps — opt-in debug-mode only.
  if (dumpBinaries) {
    const binPaths = Object.keys(zip).filter(k => k.endsWith('.bin')).sort();
    const total = binPaths.length;
    let doneBins = 0;
    for (const path of binPaths) {
      doneBins++;
      const pct = 35 + Math.round((doneBins / total) * 55);
      onProgress?.(`${path.split('/').pop()}...`, pct);
      const dump = dumpBinary(path, zip[path]);
      out.binaryDumps.push(dump);
      out.summary.fileCount++;
      out.summary.totalRecords += dump.recCount;
      await tick();
    }
  } else {
    // Still count files for summary (cheap).
    const binPaths = Object.keys(zip).filter(k => k.endsWith('.bin'));
    out.summary.fileCount += binPaths.length;
  }

  await tick();

  if (readXml) {
    const xmlPaths = Object.keys(zip).filter(k => k.endsWith('.xml') || k.endsWith('.rels'));
    for (let i = 0; i < xmlPaths.length; i++) {
      const path = xmlPaths[i];
      if (i % 5 === 0) {
        onProgress?.(`XML files...`, 92 + Math.round((i / xmlPaths.length) * 6));
        await tick();
      }
      try { out.xmlFiles[path] = dec8.decode(zip[path]); } catch { /* skip */ }
    }
  }

  (zip as any) = null;

  onProgress?.('Done', 100);
  await tick();
  return out;
}

// Re-export for tests that need internal helpers (kept private to the lib).
export { parseStyles, isDateFormatId, isDateNumFmtString };
