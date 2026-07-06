import type { BinaryDump, RawRecord } from './types.js';
import { records, readU32, dec16, hex } from './record-stream.js';

// Brt* record-dumper used for debugging / inspection of arbitrary .bin parts.
// Not part of the worksheet/sheet-data flow; ships in P5 behind an opt-in
// flag. For now it always runs (preserves the v0.2 API surface).
export function dumpBinary(path: string, data: Uint8Array, maxRec = 200): BinaryDump {
  const recordsArr: RawRecord[] = [];
  const typeSummary: Record<string, number> = {};
  let total = 0;

  for (const r of records(data)) {
    const key = '0x' + r.type.toString(16).toUpperCase().padStart(4, '0');
    typeSummary[key] = (typeSummary[key] || 0) + 1;
    total++;

    if (recordsArr.length >= maxRec) continue;

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
