import { unzipSync } from 'fflate';

const dec16 = new TextDecoder('utf-16le');
const dec8 = new TextDecoder('utf-8');

export type ProgressCallback = (msg: string, pct: number) => void;

function tick(): Promise<void> {
  return new Promise(r => setTimeout(r, 0));
}

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

const BRT_ROW_HEADER = 0x00;
const BRT_CELL_BLANK = 0x01;
const BRT_CELL_RK = 0x02;
const BRT_CELL_ERROR = 0x03;
const BRT_CELL_BOOL = 0x04;
const BRT_CELL_REAL = 0x05;
const BRT_CELL_ST = 0x06;
const BRT_CELL_ISST = 0x07;
const BRT_FMLA_STRING = 0x08;
const BRT_FMLA_NUM = 0x09;
const BRT_FMLA_BOOL = 0x0A;
const BRT_FMLA_ERROR = 0x0B;
const BRT_SHORT_BLANK = 0x0C;
const BRT_SHORT_RK = 0x0D;
const BRT_SHORT_ERROR = 0x0E;
const BRT_SHORT_BOOL = 0x0F;
const BRT_SHORT_REAL = 0x10;
const BRT_SHORT_ST = 0x11;
const BRT_SHORT_ISST = 0x12;
const BRT_SST_ITEM = 0x13;
const BRT_BUNDLE_SH = 0x9C;
const BRT_BUNDLE_SH_NEW = 0x0E01;

const ERRORS: Record<number, string> = {
  0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!',
  0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!',
  0x2A: '#N/A', 0x2B: '#GETTING_DATA',
};

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

function* records(data: Uint8Array): Generator<{ type: number; data: Uint8Array }> {
  let off = 0;
  while (off < data.length) {
    const recStart = off;
    if (off >= data.length) break;
    let t = data[off++];
    if ((t & 0x80) !== 0) {
      if (off >= data.length) {
        throw new Error(
          `Truncated .bin: record type byte at offset ${recStart} announces a second byte but only ${data.length} bytes total remain`,
        );
      }
      t = ((t & 0x7F) << 7) | data[off++];
    }
    let s = 0, sh = 0, b: number;
    do {
      if (off >= data.length) {
        throw new Error(
          `Truncated .bin: record at offset ${recStart} (type 0x${t.toString(16)}) declared size varint overruns the buffer`,
        );
      }
      b = data[off++];
      s |= (b & 0x7F) << sh;
      sh += 7;
    } while (b & 0x80);
    if (off + s > data.length) {
      throw new Error(
        `Truncated .bin: record at offset ${recStart} (type 0x${t.toString(16)}) declared size ${s} but only ${data.length - off} bytes remain`,
      );
    }
    yield { type: t, data: data.subarray(off, off + s) };
    off += s;
  }
}

function readU32(d: Uint8Array, off: number): number {
  return (d[off] | (d[off + 1] << 8) | (d[off + 2] << 16) | (d[off + 3] << 24)) >>> 0;
}

function readF64(d: Uint8Array, off: number): number {
  return new DataView(d.buffer, d.byteOffset + off, 8).getFloat64(0, true);
}

function readWideString(d: Uint8Array, off: number): string {
  const len = readU32(d, off);
  return dec16.decode(d.subarray(off + 4, off + 4 + len * 2));
}

function readRichString(d: Uint8Array, off: number): string {
  return readWideString(d, off + 1);
}

function decodeRk(rk: number): number {
  const fx100 = rk & 0x01;
  const fInt = (rk >> 1) & 0x01;
  const num = rk >>> 2;
  let val: number;
  if (fInt) {
    val = (num << 2) >> 2;
  } else {
    const buf = new ArrayBuffer(8);
    new DataView(buf).setBigUint64(0, BigInt(num) << 34n, true);
    val = new DataView(buf).getFloat64(0, true);
  }
  if (fx100) val /= 100;
  return val;
}

function hex(d: Uint8Array, max = 48): string {
  return Array.from(d.subarray(0, Math.min(max, d.length)))
    .map(b => b.toString(16).padStart(2, '0')).join(' ');
}

// ---- workbook ----

function parseWorkbook(data: Uint8Array): string[] {
  const names: string[] = [];
  for (const r of records(data)) {
    if (r.type === BRT_BUNDLE_SH_NEW && r.data.length >= 12) {
      const rIdLen = readU32(r.data, 8);
      const nameOff = 12 + rIdLen * 2;
      if (nameOff + 4 <= r.data.length) names.push(readWideString(r.data, nameOff));
    } else if (r.type === BRT_BUNDLE_SH && r.data.length >= 9) {
      const rIdLen = readU32(r.data, 5);
      const nameOff = 9 + rIdLen * 2;
      if (nameOff + 4 <= r.data.length) names.push(readWideString(r.data, nameOff));
    }
  }
  return names;
}

// ---- shared strings ----

function parseSharedStrings(data: Uint8Array): string[] {
  const list: string[] = [];
  for (const r of records(data)) {
    if (r.type === BRT_SST_ITEM && r.data.length >= 5) list.push(readRichString(r.data, 0));
  }
  return list;
}

// ---- sheets ----

function parseSheet(data: Uint8Array, ss: string[]): ParsedRow[] {
  const rows: ParsedRow[] = [];
  let curRow: ParsedRow | null = null;
  let prevCol = -1;

  for (const r of records(data)) {
    const d = r.data;
    if (r.type === BRT_ROW_HEADER && d.length >= 4) {
      curRow = { row: readU32(d, 0), cols: {} };
      rows.push(curRow);
      prevCol = -1;
      continue;
    }
    if (!curRow) continue;

    if (r.type >= BRT_CELL_BLANK && r.type <= BRT_FMLA_ERROR) {
      const col = readU32(d, 0);
      const ixf = d.length >= 6 ? (d[4] | (d[5] << 8)) | (d[5] & 0x80 ? 0xFFFF0000 : 0) : undefined;
      const cell = readCell(r.type, d, 8, ss);
      if (cell) { cell.ixf = ixf; curRow.cols[col] = cell; }
      prevCol = col;
    } else if (r.type >= BRT_SHORT_BLANK && r.type <= BRT_SHORT_ISST) {
      const col = prevCol + 1;
      const ixf = d.length >= 4 ? readU16(d, 2) : undefined;
      const cell = readShortCell(r.type, d, 4, ss);
      if (cell) { cell.ixf = ixf; curRow.cols[col] = cell; }
      prevCol = col;
    }
  }
  return rows;
}

function readCell(type: number, d: Uint8Array, off: number, ss: string[]): Cell | null {
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
      // BrtRichStr needs at least 1 flag + 4 cch bytes
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

function readShortCell(type: number, d: Uint8Array, off: number, ss: string[]): Cell | null {
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
function formatYMD(y: number, m: number, d: number): string {
  return String(y).padStart(4, '0') + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}
function readU16(d: Uint8Array, off: number): number {
  return d[off] | (d[off + 1] << 8);
}

// ---- binary dumper ----

function dumpBinary(path: string, data: Uint8Array, maxRec = 200): BinaryDump {
  const recordsArr: RawRecord[] = [];
  const typeSummary: Record<string, number> = {};
  let total = 0;

  for (const r of records(data)) {
    const key = '0x' + r.type.toString(16).toUpperCase().padStart(4, '0');
    typeSummary[key] = (typeSummary[key] || 0) + 1;
    total++;

    if (recordsArr.length >= maxRec) continue; // stop collecting after max

    const strings: string[] = [];
    for (let off = 0; off + 4 < r.data.length;) {
      const len = readU32(r.data, off);
      if (len > 0 && len < 200 && off + 4 + len * 2 <= r.data.length) {
        try {
          const s = dec16.decode(r.data.subarray(off + 4, off + 4 + len * 2)).replace(/\0/g, '');
          if (s.length >= 2) strings.push(s);
        } catch { /* skip */ }
        off += 4 + len * 2;
      } else {
        off++;
      }
    }

    recordsArr.push({
      type: key,
      typeNum: r.type,
      size: r.data.length,
      hex: hex(r.data, 32),
      strings,
    });
  }

  if (total > maxRec) {
    recordsArr.push({ type: '...', typeNum: -1, size: 0, hex: `[${total - maxRec} more records omitted]`, strings: [] });
  }

  return { path, size: data.length, recCount: total, records: recordsArr, typeSummary };
}

// ---- pivot cache parser ----

function parsePivotCache(name: string, def: Uint8Array, recs: Uint8Array): PivotCacheTable {
  const fieldNames: string[] = [];
  const sharedItems: (string | null)[][] = [];
  let curItems: (string | null)[] = [];
  let fallbackItems: (string | null)[] = [];
  let hierItems: (string | null)[] = [];
  let has1F81 = false;

  function pushField() {
    if (fieldNames.length === 0) return;
    const src = has1F81 ? curItems : (curItems.length > 0 ? curItems : (fallbackItems.length > 0 ? fallbackItems : (hierItems.length > 0 ? hierItems : curItems)));
    sharedItems.push(src);
  }

  for (const r of records(def)) {
    if (r.type === 0x1B81) {
      pushField();
      curItems = [];
      fallbackItems = [];
      hierItems = [];
      has1F81 = false;
      const d = r.data;
      if (d.length >= 24) {
        const nameLen = readU32(d, 20);
        if (nameLen > 0 && nameLen < 100 && 24 + nameLen * 2 <= d.length) {
          fieldNames.push(dec16.decode(d.subarray(24, 24 + nameLen * 2)));
        }
      }
    } else if (r.type === 0x001F) {
      // 0x001F = shared items fallback (for fields without 0x1F81)
      // Format: uint32(len) + UTF-16LE string
      const d = r.data;
      if (d.length >= 4) {
        const slen = readU32(d, 0);
        if (slen > 0 && slen < 500) {
          const endOff = 4 + slen * 2;
          if (endOff <= d.length) {
            fallbackItems.push(dec16.decode(d.subarray(4, endOff)));
          }
        }
      }
    } else if (r.type === 0x18 && r.data.length > 4) {
      // 0x18 = field member/group items (fallback for fields without 0x001F/0x1F81)
      const len = readU32(r.data, 0);
      if (len > 0 && len < 200 && 4 + len * 2 <= r.data.length) {
        hierItems.push(dec16.decode(r.data.subarray(4, 4 + len * 2)));
      }
    } else if (r.type === 0x1F81 && r.data.length >= 6 && (r.data[0] === 0x20 || r.data[0] === 0x02)) {
      has1F81 = true;
      const d = r.data;
      if (d[0] === 0x20) {
        const count = readU32(d, 2);
        for (let i = 0, off = 6; i < count && off + 8 <= d.length; i++, off += 8) {
          curItems.push(formatYMD(d[off] | (d[off + 1] << 8), d[off + 2] | (d[off + 3] << 8), d[off + 4] | (d[off + 5] << 8)));
        }
      } else if (d[0] === 0x02) {
        const count = readU32(d, 2);
        for (let i = 0, off = 6; i < count && off + 4 <= d.length; i++) {
          const slen = readU32(d, off);
          if (slen > 0 && slen < 500 && off + 4 + slen * 2 <= d.length) {
            curItems.push(dec16.decode(d.subarray(off + 4, off + 4 + slen * 2)));
            off += 4 + slen * 2;
          } else break;
        }
      }
    } else if (r.type === 0x0020 && r.data.length >= 8) {
      const d = r.data;
      curItems.push(formatYMD(d[0] | (d[1] << 8), d[2] | (d[3] << 8), d[4] | (d[5] << 8)));
    }
  }
  pushField();

  // Sample record bodies for type detection
  const recBodies: Uint8Array[] = [];
  for (const r of records(recs)) {
    if (r.type === 0x2101) break;
    if (r.type !== 0x0021) continue;
    recBodies.push(r.data);
    if (recBodies.length >= 50) break;
  }

  const fieldCount = fieldNames.length;
  const fields: ('str' | 'f64' | 'u32' | null)[] = new Array(fieldCount).fill(null);
  const recOffsets = new Int32Array(recBodies.length);

  for (let fi = 0; fi < fieldCount; fi++) {
    for (let ri = 0; ri < recBodies.length; ri++) {
      const body = recBodies[ri];
      let off = recOffsets[ri];
      if (off < 0 || off + 4 > body.length) continue;
      const len = readU32(body, off);
      if (len >= 3 && len < 200 && off + 4 + len * 2 <= body.length) {
        try {
          const s = dec16.decode(body.subarray(off + 4, off + 4 + len * 2));
          let valid = true;
          let alpha = 0;
          for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 32 && c !== 10 && c !== 13) { valid = false; break; }
            if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122)) alpha++;
          }
          if (valid && alpha > 0) { fields[fi] = 'str'; break; }
        } catch {}
      }
    }
    if (fields[fi] === 'str') {
      for (let ri = 0; ri < recBodies.length; ri++) {
        const body = recBodies[ri];
        let off = recOffsets[ri];
        if (off >= 0) {
          const len = readU32(body, off);
          recOffsets[ri] = off + 4 + len * 2;
        }
      }
      continue;
    }

    let isF64 = false;
    for (let ri = 0; ri < recBodies.length; ri++) {
      const body = recBodies[ri];
      let off = recOffsets[ri];
      if (off < 0 || off + 8 > body.length) continue;
      const f = readF64(body, off);
      const lo = readU32(body, off);
      const hi = readU32(body, off + 4);
      if (isFinite(f) && !isNaN(f) && (lo !== 0 || hi !== 0) &&
          Math.abs(f) > 1e-10 && Math.abs(f) < 1e20 &&
          !(hi === 0 && lo < 100000)) {
        isF64 = true;
        break;
      }
    }
    fields[fi] = isF64 ? 'f64' : 'u32';
    const sz = isF64 ? 8 : 4;
    for (let ri = 0; ri < recBodies.length; ri++) {
      if (recOffsets[ri] >= 0) recOffsets[ri] += sz;
    }
  }

  // Post-process: fill f64 gaps, trailing deficit
  for (let i = 1; i < fields.length - 1; i++) {
    if (fields[i] === 'u32' && fields[i - 1] === 'f64' && fields[i + 1] === 'f64') fields[i] = 'f64';
  }
  let changed = true;
  while (changed) {
    changed = false;
    for (let i = 1; i < fields.length - 1; i++) {
      if (fields[i] === 'u32' && fields[i - 1] === 'f64' && fields[i + 1] === 'f64') { fields[i] = 'f64'; changed = true; }
    }
  }

  let lastStr = -1;
  for (let i = 0; i < fields.length; i++) { if (fields[i] === 'str') lastStr = i; }

  for (const body of recBodies) {
    let consumed = 0;
    for (let i = 0; i < fields.length; i++) {
      if (fields[i] === 'str') {
        if (consumed + 4 > body.length) { consumed = -1; break; }
        const len = readU32(body, consumed);
        consumed += (len >= 3 && consumed + 4 + len * 2 <= body.length) ? 4 + len * 2 : 4;
      } else { consumed += fields[i] === 'f64' ? 8 : 4; }
    }
    if (consumed < 0) continue;
    const deficit = body.length - consumed;
    if (deficit > 0 && deficit % 4 === 0) {
      const mis = deficit / 4;
      let converted = 0;
      for (let i = fields.length - 1; i > lastStr && converted < mis; i--) {
        if (fields[i] === 'u32') { fields[i] = 'f64'; converted++; }
      }
    }
  }

  // Trim field count to match actual record storage
  if (recBodies.length > 0) {
    let best = fields.length;
    for (let check = fields.length; check >= 1; check--) {
      let allFit = true;
      for (const body of recBodies) {
        let off = 0;
        for (let i = 0; i < check && off <= body.length; i++) {
          if (fields[i] === 'str') {
            if (off + 4 > body.length) { allFit = false; break; }
            const len = readU32(body, off);
            off += (len >= 3 && off + 4 + len * 2 <= body.length) ? 4 + len * 2 : 4;
          } else { off += fields[i] === 'f64' ? 8 : 4; }
        }
        if (off > body.length) { allFit = false; break; }
      }
      if (allFit) { best = check; break; }
    }
    if (best < fields.length) fields.length = best;
  }

  // Decode records, resolving shared-item indexes via 0x001F data
  const rows: (string | number | null)[][] = [];
  for (const r of records(recs)) {
    if (r.type === 0x2101) break;
    if (r.type !== 0x0021) continue;

    const d = r.data;
    const values: (string | number | null)[] = [];
    let off = 0;
    let fi = 0;

    while (off < d.length && fi < fields.length) {
      const ft = fields[fi];
      if (ft === 'str') {
        const len = readU32(d, off);
        if (len >= 3 && off + 4 + len * 2 <= d.length) {
          values.push(dec16.decode(d.subarray(off + 4, off + 4 + len * 2)));
          off += 4 + len * 2;
        } else { values.push(null); off += 4; }
      } else if (ft === 'f64') {
        if (off + 8 <= d.length) {
          const f = readF64(d, off);
          values.push(f === 0 ? 0 : (Math.abs(f) >= 1 ? parseFloat(f.toFixed(4)) : parseFloat(f.toFixed(8))));
          off += 8;
        } else { values.push(null); off += 4; }
      } else {
        // u32 — resolve shared-item index if field has shared items
        if (off + 4 <= d.length) {
          const v = readU32(d, off);
          const si = fi < sharedItems.length ? sharedItems[fi] : null;
          if (si && si.length > 0 && v < si.length && si[v] !== null) {
            values.push(si[v]);
          } else {
            values.push(v);
          }
        } else values.push(null);
        off += 4;
      }
      fi++;
    }

    rows.push(values);
  }

  // Trim ghost fields — fields with no data in any of the first few rows
  if (rows.length > 0) {
    let lastPopulated = 0;
    for (let ri = 0; ri < Math.min(5, rows.length); ri++) {
      for (let fi = rows[ri].length - 1; fi >= 0; fi--) {
        if (rows[ri][fi] !== undefined && fi > lastPopulated) lastPopulated = fi;
      }
    }
    if (lastPopulated + 1 < fieldNames.length) fieldNames.length = lastPopulated + 1;
  }

  return { name, fieldNames, rows, rowCount: recBodies.length > 50 ? recBodies.length : rows.length };
}
