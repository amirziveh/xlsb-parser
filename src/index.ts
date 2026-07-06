import { unzipSync } from 'fflate';
import type {
  ProgressCallback, ParsedXlsb,
} from './types.js';
import {
  dec8, records, readU32, readWideString,
  BRT_BUNDLE_SH, BRT_BUNDLE_SH_NEW,
} from './record-stream.js';
import { parseWorkbook } from './workbook.js';
import { parseSharedStrings } from './shared-strings.js';
import { parseSheet } from './sheet.js';
import { dumpBinary } from './dump.js';
import { parsePivotCache } from './pivot-cache.js';

// Re-export the public types so consumers can `import type { ... } from 'xlsb-parser'`.
export type {
  ProgressCallback, Cell, ParsedRow, Sheet, RawRecord, BinaryDump,
  PivotCacheTable, ParsedXlsb,
} from './types.js';

function tick(): Promise<void> {
  return new Promise(r => setTimeout(r, 0));
}

export async function parseXlsb(
  data: ArrayBuffer | Uint8Array,
  onProgress?: ProgressCallback,
): Promise<ParsedXlsb> {
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

  for (let i = 0; i < sheetNames.length; i++) {
    const key = `xl/worksheets/sheet${i + 1}.bin`;
    const sd = zip[key];
    if (sd) {
      onProgress?.(`Sheet "${sheetNames[i]}"...`, 15 + Math.round((i / sheetNames.length) * 20));
      const rows = parseSheet(sd, out.sharedStrings);
      const totalCells = rows.reduce((a, r) => a + Object.keys(r.cols).length, 0);
      out.sheets.push({ name: sheetNames[i], rows, totalCells });
      await tick();
    }
  }

  // Parse pivot caches
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

  await tick();

  const xmlPaths = Object.keys(zip).filter(k => k.endsWith('.xml') || k.endsWith('.rels'));
  for (let i = 0; i < xmlPaths.length; i++) {
    const path = xmlPaths[i];
    if (i % 5 === 0) {
      onProgress?.(`XML files...`, 92 + Math.round((i / xmlPaths.length) * 6));
      await tick();
    }
    try { out.xmlFiles[path] = dec8.decode(zip[path]); } catch { /* skip */ }
  }

  (zip as any) = null;

  onProgress?.('Done', 100);
  await tick();
  return out;
}
