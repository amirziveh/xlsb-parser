import type { PivotCacheTable } from './types.js';
import { records, readU32, readF64, dec16 } from './record-stream.js';

function formatYMD(y: number, m: number, d: number): string {
  return (
    String(y).padStart(4, '0') + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0')
  );
}

// Parse xl/pivotCache/pivotCacheDefinitionN.bin + matching records part.
//
// NOTE: This module decodes pivot-cache records using a heuristic field-type
// detector (sample-then-probe). P3 of the world-level roadmap replaces this
// with a spec-driven sequential per-field reader (MS-ODFFXML §3). Until
// then, the decoded rows for non-trivial pivot caches may be wrong on
// field boundaries. The output is best-effort and the fieldNames list
// is trustworthy; rows should be treated as approximations.
export function parsePivotCache(name: string, def: Uint8Array, recs: Uint8Array): PivotCacheTable {
  const fieldNames: string[] = [];
  const sharedItems: (string | null)[][] = [];
  let curItems: (string | null)[] = [];
  let fallbackItems: (string | null)[] = [];
  let hierItems: (string | null)[] = [];
  let has1F81 = false;

  function pushField() {
    if (fieldNames.length === 0) return;
    const src = has1F81
      ? curItems
      : curItems.length > 0
        ? curItems
        : fallbackItems.length > 0
          ? fallbackItems
          : hierItems.length > 0
            ? hierItems
            : curItems;
    sharedItems.push(src);
  }

  for (const r of records(def)) {
    if (r.type === 0x1b81) {
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
    } else if (r.type === 0x001f) {
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
      const len = readU32(r.data, 0);
      if (len > 0 && len < 200 && 4 + len * 2 <= r.data.length) {
        hierItems.push(dec16.decode(r.data.subarray(4, 4 + len * 2)));
      }
    } else if (
      r.type === 0x1f81 &&
      r.data.length >= 6 &&
      (r.data[0] === 0x20 || r.data[0] === 0x02)
    ) {
      has1F81 = true;
      const d = r.data;
      if (d[0] === 0x20) {
        const count = readU32(d, 2);
        for (let i = 0, off = 6; i < count && off + 8 <= d.length; i++, off += 8) {
          curItems.push(
            formatYMD(
              d[off] | (d[off + 1] << 8),
              d[off + 2] | (d[off + 3] << 8),
              d[off + 4] | (d[off + 5] << 8),
            ),
          );
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
      const off = recOffsets[ri];
      if (off < 0 || off + 4 > body.length) continue;
      const len = readU32(body, off);
      if (len >= 3 && len < 200 && off + 4 + len * 2 <= body.length) {
        try {
          const s = dec16.decode(body.subarray(off + 4, off + 4 + len * 2));
          let valid = true;
          let alpha = 0;
          for (let i = 0; i < s.length; i++) {
            const c = s.charCodeAt(i);
            if (c < 32 && c !== 10 && c !== 13) {
              valid = false;
              break;
            }
            if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122)) alpha++;
          }
          if (valid && alpha > 0) {
            fields[fi] = 'str';
            break;
          }
        } catch {
          /* malformed UTF-16; skip this candidate */
        }
      }
    }
    if (fields[fi] === 'str') {
      for (let ri = 0; ri < recBodies.length; ri++) {
        const body = recBodies[ri];
        const off = recOffsets[ri];
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
      const off = recOffsets[ri];
      if (off < 0 || off + 8 > body.length) continue;
      const f = readF64(body, off);
      const lo = readU32(body, off);
      const hi = readU32(body, off + 4);
      if (
        isFinite(f) &&
        !isNaN(f) &&
        (lo !== 0 || hi !== 0) &&
        Math.abs(f) > 1e-10 &&
        Math.abs(f) < 1e20 &&
        !(hi === 0 && lo < 100000)
      ) {
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

  for (let i = 1; i < fields.length - 1; i++) {
    if (fields[i] === 'u32' && fields[i - 1] === 'f64' && fields[i + 1] === 'f64')
      fields[i] = 'f64';
  }
  let changed = true;
  while (changed) {
    changed = false;
    for (let i = 1; i < fields.length - 1; i++) {
      if (fields[i] === 'u32' && fields[i - 1] === 'f64' && fields[i + 1] === 'f64') {
        fields[i] = 'f64';
        changed = true;
      }
    }
  }

  let lastStr = -1;
  for (let i = 0; i < fields.length; i++) {
    if (fields[i] === 'str') lastStr = i;
  }

  for (const body of recBodies) {
    let consumed = 0;
    for (let i = 0; i < fields.length; i++) {
      if (fields[i] === 'str') {
        if (consumed + 4 > body.length) {
          consumed = -1;
          break;
        }
        const len = readU32(body, consumed);
        consumed += len >= 3 && consumed + 4 + len * 2 <= body.length ? 4 + len * 2 : 4;
      } else {
        consumed += fields[i] === 'f64' ? 8 : 4;
      }
    }
    if (consumed < 0) continue;
    const deficit = body.length - consumed;
    if (deficit > 0 && deficit % 4 === 0) {
      const mis = deficit / 4;
      let converted = 0;
      for (let i = fields.length - 1; i > lastStr && converted < mis; i--) {
        if (fields[i] === 'u32') {
          fields[i] = 'f64';
          converted++;
        }
      }
    }
  }

  if (recBodies.length > 0) {
    let best = fields.length;
    for (let check = fields.length; check >= 1; check--) {
      let allFit = true;
      for (const body of recBodies) {
        let off = 0;
        for (let i = 0; i < check && off <= body.length; i++) {
          if (fields[i] === 'str') {
            if (off + 4 > body.length) {
              allFit = false;
              break;
            }
            const len = readU32(body, off);
            off += len >= 3 && off + 4 + len * 2 <= body.length ? 4 + len * 2 : 4;
          } else {
            off += fields[i] === 'f64' ? 8 : 4;
          }
        }
        if (off > body.length) {
          allFit = false;
          break;
        }
      }
      if (allFit) {
        best = check;
        break;
      }
    }
    if (best < fields.length) fields.length = best;
  }

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
        } else {
          values.push(null);
          off += 4;
        }
      } else if (ft === 'f64') {
        if (off + 8 <= d.length) {
          const f = readF64(d, off);
          values.push(
            f === 0 ? 0 : Math.abs(f) >= 1 ? parseFloat(f.toFixed(4)) : parseFloat(f.toFixed(8)),
          );
          off += 8;
        } else {
          values.push(null);
          off += 4;
        }
      } else {
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

  if (rows.length > 0) {
    let lastPopulated = 0;
    for (let ri = 0; ri < Math.min(5, rows.length); ri++) {
      for (let fi = rows[ri].length - 1; fi >= 0; fi--) {
        if (rows[ri][fi] !== undefined && fi > lastPopulated) lastPopulated = fi;
      }
    }
    if (lastPopulated + 1 < fieldNames.length) fieldNames.length = lastPopulated + 1;
  }

  return {
    name,
    fieldNames,
    rows,
    rowCount: recBodies.length > 50 ? recBodies.length : rows.length,
  };
}
