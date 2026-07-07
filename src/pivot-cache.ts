import type { PivotCacheTable, PivotCacheField, PivotCacheCell } from './types.js';
import {
  records, readU32, readF64, dec16,
  BRT_BEGIN_PCD_FIELD, BRT_BEGIN_PCD_ATBL, BRT_BEGIN_PCDIRUN,
  BRT_PCDI_STRING, BRT_PCDI_STRING2,
  BRT_PCDIDATETIME, BRT_PCDINUMBER,
  BRT_PCDIBOOLEAN, BRT_PCDIERROR, BRT_PCDIMISSING,
  BRT_BEGIN_PIVOT_CACHE_RECORDS, BRT_PC_RECORD, BRT_PC_RECORD_DT,
  BRT_END_PIVOT_CACHE_RECORDS,
} from './record-stream.js';

function decodePCDIDateTime(d: Uint8Array, off: number): string {
  const yr = d[off] | (d[off + 1] << 8);
  const mon = d[off + 2] | (d[off + 3] << 8);
  const dom = d[off + 4];
  const hr = d[off + 5];
  const min = d[off + 6];
  const sec = d[off + 7];
  const date = String(yr).padStart(4, '0') + '-' +
    String(mon).padStart(2, '0') + '-' + String(dom).padStart(2, '0');
  if (hr === 0 && min === 0 && sec === 0) return date;
  const t = String(hr).padStart(2, '0') + ':' +
    String(min).padStart(2, '0') + ':' + String(sec).padStart(2, '0');
  return date + 'T' + t;
}

interface FieldBuilder {
  name: string;
  isSrc: boolean;
  fNumField: boolean;
  fDateInField: boolean;
  fHasTextItem: boolean;
  kind: PivotCacheField['kind'];
  sharedItems: (PivotCacheCell | null)[];
}

function parseDefinition(def: Uint8Array): FieldBuilder[] {
  const fields: FieldBuilder[] = [];
  let cur: FieldBuilder | null = null;

  function closeField() {
    if (cur) {
      if (cur.sharedItems.length > 0 && cur.kind === 'string') cur.kind = 'indexed';
      fields.push(cur);
      cur = null;
    }
  }

  for (const r of records(def)) {
    if (r.type === BRT_BEGIN_PCD_FIELD) {
      closeField();
      const d = r.data;
      const isSrc = d.length >= 1 ? (d[0] & 0x04) !== 0 : false;
      let name = '';
      if (d.length >= 24) {
        const nameLen = readU32(d, 20);
        if (nameLen > 0 && nameLen < 32768 && 24 + nameLen * 2 <= d.length) {
          name = dec16.decode(d.subarray(24, 24 + nameLen * 2));
        }
      }
      cur = {
        name, isSrc, fNumField: false, fDateInField: false, fHasTextItem: false,
        kind: 'string', sharedItems: [],
      };
    } else if (r.type === BRT_BEGIN_PCD_ATBL && cur) {
      const d = r.data;
      const flags = d.length >= 2 ? d[0] | (d[1] << 8) : 0;
      cur.fNumField = (flags & 0x40) !== 0;
      cur.fDateInField = (flags & 0x04) !== 0;
      cur.fHasTextItem = (flags & 0x08) !== 0;
      if (cur.fNumField) cur.kind = 'number';
      else if (cur.fDateInField && !cur.fHasTextItem) cur.kind = 'date';
      else cur.kind = 'string';
    } else if (cur) {
      const d = r.data;
      if (r.type === BRT_PCDI_STRING || r.type === BRT_PCDI_STRING2) {
        if (d.length >= 4) {
          const slen = readU32(d, 0);
          if (slen > 0 && slen < 32768 && 4 + slen * 2 <= d.length) {
            cur.sharedItems.push({ t: 's', v: dec16.decode(d.subarray(4, 4 + slen * 2)) });
          }
        }
      } else if (r.type === BRT_PCDIDATETIME) {
        if (d.length >= 8) cur.sharedItems.push({ t: 'd', v: decodePCDIDateTime(d, 0) });
      } else if (r.type === BRT_PCDINUMBER) {
        if (d.length >= 8) cur.sharedItems.push({ t: 'n', v: readF64(d, 0) });
      } else if (r.type === BRT_PCDIBOOLEAN) {
        if (d.length >= 1) cur.sharedItems.push({ t: 'b', v: d[0] !== 0 });
      } else if (r.type === BRT_PCDIERROR) {
        if (d.length >= 1) cur.sharedItems.push({ t: 'e', v: errorName(d[0]) });
      } else if (r.type === BRT_PCDIMISSING) {
        cur.sharedItems.push({ t: 'blank' });
      } else if (r.type === BRT_BEGIN_PCDIRUN) {
        decodeRun(d, cur);
      }
    }
  }
  closeField();
  return fields;
}

function decodeRun(d: Uint8Array, cur: FieldBuilder) {
  if (d.length < 6) return;
  const mdx = d[0] | (d[1] << 8);
  const citems = readU32(d, 2);
  let off = 6;
  for (let i = 0; i < citems && off + 1 <= d.length; i++) {
    if (mdx === 0x02) {
      if (off + 4 > d.length) break;
      const slen = readU32(d, off); off += 4;
      if (slen > 0 && slen < 32768 && off + slen * 2 <= d.length) {
        cur.sharedItems.push({ t: 's', v: dec16.decode(d.subarray(off, off + slen * 2)) });
        off += slen * 2;
      } else break;
    } else if (mdx === 0x01) {
      if (off + 8 > d.length) break;
      cur.sharedItems.push({ t: 'n', v: readF64(d, off) }); off += 8;
    } else if (mdx === 0x10) {
      if (off + 1 > d.length) break;
      cur.sharedItems.push({ t: 'e', v: errorName(d[off]) }); off += 1;
    } else if (mdx === 0x20) {
      if (off + 8 > d.length) break;
      cur.sharedItems.push({ t: 'd', v: decodePCDIDateTime(d, off) }); off += 8;
    } else break;
  }
}

function errorName(code: number): string {
  switch (code) {
    case 0x00: return '#NULL!';
    case 0x07: return '#DIV/0!';
    case 0x0f: return '#VALUE!';
    case 0x17: return '#REF!';
    case 0x1d: return '#NAME?';
    case 0x24: return '#NUM!';
    case 0x2a: return '#N/A';
    default: return '#ERR!';
  }
}

export function parsePivotCache(name: string, def: Uint8Array, recs: Uint8Array): PivotCacheTable {
  const builders = parseDefinition(def);
  const fields: PivotCacheField[] = builders.map(b => ({
    name: b.name, isSrc: b.isSrc, kind: b.kind, sharedItems: b.sharedItems,
  }));
  const fieldNames = builders.map(b => b.name);
  const srcFields = builders.filter(b => b.isSrc);

  const rows: PivotCacheCell[][] = [];
  let rowCount = 0;
  let recordCount = 0;
  let dtRow: PivotCacheCell[] | null = null;
  let dtIdx = 0;

  for (const r of records(recs)) {
    if (r.type === BRT_BEGIN_PIVOT_CACHE_RECORDS) {
      if (r.data.length >= 4) rowCount = readU32(r.data, 0);
      continue;
    }
    if (r.type === BRT_END_PIVOT_CACHE_RECORDS) break;
    if (r.type === BRT_PC_RECORD) {
      dtRow = null;
      const row = decodeRgbRow(r.data, srcFields);
      if (row) { rows.push(row); recordCount++; }
    } else if (r.type === BRT_PC_RECORD_DT) {
      dtRow = [];
      dtIdx = 0;
    } else if (dtRow) {
      const cell = decodeDtCell(r, srcFields[dtIdx]);
      if (cell) dtRow.push(cell);
      dtIdx++;
      if (dtIdx >= srcFields.length) {
        rows.push(dtRow); recordCount++; dtRow = null;
      }
    }
  }

  return { name, fieldNames, fields, rows, rowCount, recordCount };
}

function decodeRgbRow(d: Uint8Array, srcFields: FieldBuilder[]): PivotCacheCell[] | null {
  const out: PivotCacheCell[] = [];
  let off = 0;
  for (const f of srcFields) {
    if (f.kind === 'indexed') {
      if (off + 4 > d.length) { out.push({ t: 'blank' }); continue; }
      const idx = readU32(d, off); off += 4;
      out.push(f.sharedItems[idx] ?? { t: 'blank' });
    } else if (f.kind === 'number') {
      if (off + 8 > d.length) { out.push({ t: 'blank' }); continue; }
      out.push({ t: 'n', v: readF64(d, off) }); off += 8;
    } else if (f.kind === 'date') {
      if (off + 8 > d.length) { out.push({ t: 'blank' }); continue; }
      out.push({ t: 'd', v: decodePCDIDateTime(d, off) }); off += 8;
    } else if (f.kind === 'string') {
      if (off + 4 > d.length) { out.push({ t: 'blank' }); continue; }
      const slen = readU32(d, off);
      if (slen > 0 && slen < 32768 && off + 4 + slen * 2 <= d.length) {
        out.push({ t: 's', v: dec16.decode(d.subarray(off + 4, off + 4 + slen * 2)) });
        off += 4 + slen * 2;
      } else { out.push({ t: 'blank' }); off += 4; }
    } else {
      if (off + 4 > d.length) { out.push({ t: 'blank' }); continue; }
      out.push({ t: 'blank' }); off += 4;
    }
  }
  return out;
}

function decodeDtCell(r: { type: number; data: Uint8Array }, f: FieldBuilder | undefined): PivotCacheCell | null {
  if (!f) return { t: 'blank' };
  const d = r.data;
  if (r.type === BRT_PCDI_STRING || r.type === BRT_PCDI_STRING2) {
    if (d.length >= 4) {
      const slen = readU32(d, 0);
      if (slen > 0 && slen < 32768 && 4 + slen * 2 <= d.length)
        return { t: 's', v: dec16.decode(d.subarray(4, 4 + slen * 2)) };
    }
    return { t: 'blank' };
  }
  if (r.type === BRT_PCDINUMBER) return d.length >= 8 ? { t: 'n', v: readF64(d, 0) } : { t: 'blank' };
  if (r.type === BRT_PCDIDATETIME) return d.length >= 8 ? { t: 'd', v: decodePCDIDateTime(d, 0) } : { t: 'blank' };
  if (r.type === BRT_PCDIBOOLEAN) return d.length >= 1 ? { t: 'b', v: d[0] !== 0 } : { t: 'blank' };
  if (r.type === BRT_PCDIERROR) return d.length >= 1 ? { t: 'e', v: errorName(d[0]) } : { t: 'blank' };
  if (r.type === BRT_PCDIMISSING) return { t: 'blank' };
  return { t: 'blank' };
}

/** Exported for streaming use by `openXlsb`. */
export { parseDefinition };

export function* streamPivotRows(def: Uint8Array, recs: Uint8Array): Generator<PivotCacheCell[]> {
  const builders = parseDefinition(def);
  const srcFields = builders.filter(b => b.isSrc);
  let dtRow: PivotCacheCell[] | null = null;
  let dtIdx = 0;
  for (const r of records(recs)) {
    if (r.type === BRT_BEGIN_PIVOT_CACHE_RECORDS) continue;
    if (r.type === BRT_END_PIVOT_CACHE_RECORDS) break;
    if (r.type === BRT_PC_RECORD) {
      dtRow = null;
      const row = decodeRgbRow(r.data, srcFields);
      if (row) yield row;
    } else if (r.type === BRT_PC_RECORD_DT) {
      dtRow = []; dtIdx = 0;
    } else if (dtRow) {
      const cell = decodeDtCell(r, srcFields[dtIdx]);
      if (cell) dtRow.push(cell);
      dtIdx++;
      if (dtIdx >= srcFields.length) { yield dtRow; dtRow = null; }
    }
  }
}
