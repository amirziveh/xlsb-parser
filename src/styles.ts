// xl/styles.bin reader + date-format detection heuristics.
//
// XLSB styles.bin layout per MS-XLSB §2.1.7: a sequence of records starting
// with BrtFmt (numFmt definitions), BrtCellXF/XF records (cell format styles),
// and so on. We only need:
//   - BrtFmt (0x02D9?): gives numFmtId + format string. Stored in numFmts map.
//   - BrtCellXF (0x01F9?): gives one cellXfs entry whose numFmtId is at byte 0.
// Only numFmtId (first 2 bytes of the XF record) is needed; we ignore the
// rest of the cell formatting (font, borders, alignment).
//
// Date detection: built-in numeric format IDs that are dates per Excel:
//   14-22, 27-36, 45-47, 50-58, 78-81.
// Plus custom format strings containing date tokens (y/m/d/h/s) outside
// escaped/literal sections. We use a conservative heuristic: if the format
// string contains any of [dyYhHsS] or 'AM'/'PM' AND doesn't look purely
// numeric. The Excel serial → JS Date conversion uses the 1900 epoch
// (Excel incorrectly treats 1900 as a leap year; the standard offset is
// Date(1899, 11, 30) + serial * 86400000ms which both SheetJS and ExcelJS
// agree on).

import type { StylesTable } from './types.js';
import { records, readU16 } from './record-stream.js';

// BrtFmt records define custom number formats: { numFmtId, formatString }.
// BrtFmt = 0x02D9 (sometimes seen as 0x2D9 — same thing)
const BRT_FMT = 0x02d9;
// BrtCellXF (aka BrtXF in modern docs) = 0x01F9 — represents a cellXfs entry.
// Per MS-OFFBFISO §2.4.300 BrtXF: data starts with uint16 numFmtId.
const BRT_CELL_XF = 0x01f9;
// BrtStyle records also exist but we skip them.

const BUILTIN_DATE_FMT_IDS = new Set<number>([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
  52, 53, 54, 55, 56, 57, 58, 78, 79, 80, 81,
]);

// Excel's earliest author-supported date serial (1.0 = 1900-01-01 in their
// flawed calendar). JS epoch for the conversion is 1899-12-30.
const SERIAL_EPOCH_MS = Date.UTC(1899, 11, 30);

export function parseStyles(data: Uint8Array): StylesTable {
  const cellXfs: number[] = [];
  const numFmts = new Map<number, string>();

  for (const r of records(data)) {
    if (r.type === BRT_FMT && r.data.length >= 4) {
      const numFmtId = readU16(r.data, 0);
      // Format string follows: it's a wide string at offset 2.
      try {
        const len = readU32At(r.data, 2);
        if (len > 0 && len < 1000 && 6 + len * 2 <= r.data.length) {
          const s = new TextDecoder('utf-16le').decode(r.data.subarray(6, 6 + len * 2));
          numFmts.set(numFmtId, s);
        }
      } catch {
        /* skip malformed */
      }
    } else if (r.type === BRT_CELL_XF && r.data.length >= 2) {
      cellXfs.push(readU16(r.data, 0));
    }
  }

  return { cellXfs, numFmts };
}

// Local helper — readU32 isn't used here normally but we need a 4-byte read
// for the format string length field. Reuse record-stream's readU32.
import { readU32 } from './record-stream.js';
function readU32At(d: Uint8Array, off: number): number {
  return readU32(d, off);
}

// Built-in date format IDs cover the standard Excel date/time/percent-ish
// formats (Microsoft's documented table in MS-OFFBFISO §2.4.326).
export function isDateFormatId(numFmtId: number): boolean {
  return BUILTIN_DATE_FMT_IDS.has(numFmtId);
}

// Heuristic for custom numFmt strings: scan for any of d/m/y/h/s (case-
// insensitive) or the literals 'AM'/'PM' that aren't escaped (inside `\` or
// quotation marks). Conservative — prefers false negatives over false
// positives since mis-tagging non-dates as dates is worse for callers.
export function isDateNumFmtString(fmt: string): boolean {
  let escaped = false;
  let quote = false;
  for (let i = 0; i < fmt.length; i++) {
    const c = fmt[i];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (c === '\\') {
      escaped = true;
      continue;
    }
    if (c === '"') {
      quote = !quote;
      continue;
    }
    if (quote) continue;
    if (/[dmyYhHsS]/.test(c)) return true;
    // AM/PM marker — only outside quotes/escapes
    if ((c === 'A' || c === 'a') && fmt.slice(i, i + 2).match(/am|AM/)) return true;
    if ((c === 'P' || c === 'p') && fmt.slice(i, i + 2).match(/pm|PM/)) return true;
  }
  return false;
}

// Convert an Excel date serial to an ISO 8601 string using the 1900 epoch.
// `serial` can be a fraction (e.g., 0.5 = noon on day 0). Output format:
// 'YYYY-MM-DDTHH:mm:ss.sssZ' (UTC).
export function numFmtIdToSerialConvert(serial: number): string {
  const ms = SERIAL_EPOCH_MS + serial * 86400000;
  return new Date(ms).toISOString();
}
