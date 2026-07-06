import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import { buildXlsb, rowHeader, rec, u32, u16le, concat } from './helpers';

// §2.15: BRT_FMLA_STRING used `d.length >= off + 6 ? readWideString(d, off) : ''`
// as its length guard. readWideString needs only 4 bytes for the `cch` uint32,
// so the threshold `off + 6` was wrong — a record恰好 4 or 5 bytes from `off`
// would silently return '' when it actually contained a valid string. The
// bounds-check fix (§2.3) corrected this to `off + 4` and returns null below
// that. This test locks in the corrected threshold.
describe('BRT_FMLA_STRING minimum-length guard (§2.15)', () => {
  it('decodes a BrtFmlaString whose only payload is the cch uint32 (4 bytes, empty string)', async () => {
    // BrtFmlaString (0x08): col(4) + iStyleRef+reserved(4) + cch(4) + chars
    // Minimum data length: 4 + 4 + 4 = 12 bytes total → string payload 4 bytes.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), rec(0x08, concat(u32(0), new Uint8Array(4), u32(0))))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]?.t).toBe('s');
    expect(wb.sheets[0].rows[0].cols[0]?.v).toBe('');
  });

  it('decodes a BrtFmlaString with an actual value', async () => {
    const s = 'hi';
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), rec(0x08, concat(u32(0), new Uint8Array(4), u32(s.length), u16le(s))))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]?.v).toBe('hi');
  });

  it('returns null when BrtFmlaString lacks even the cch uint32', async () => {
    // 4 col + 4 iStyleRef/reserved + 2 stray bytes — no room for cch.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), rec(0x08, concat(u32(0), new Uint8Array([0, 0, 0, 0]), new Uint8Array([0x99, 0x99]))))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]).toBeUndefined();
  });
});
