import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import { buildXlsb, rowHeader, cellRealStyled, concat } from './helpers';

// §2.1: iStyleRef sign extension checked the wrong bit.
// iStyleRef is treated as a 16-bit value stored at d[4..5]; sign bit = bit 15
// (= high bit of HIGH byte = d[5]&0x80). The buggy code checked d[4]&0x80,
// which is bit 7 of the LOW byte — a positive value like 0x0080=128 got
// wrongly sign-extended to -128.
describe('iStyleRef (ixf) sign extension', () => {
  it('does NOT sign-extend when only the LOW byte has bit 7 set (value 128)', async () => {
    // d[4]=0x80 d[5]=0x00 → unsigned 128, sign bit (bit 15) is clear → expect +128
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealStyled(0, 99.0, 0x80, 0x00))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]?.ixf).toBe(128);
  });

  it('sign-extends when the HIGH byte has bit 7 set (value -1)', async () => {
    // d[4]=0xFF d[5]=0xFF → unsigned 0xFFFF, sign bit set → expect -1
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealStyled(0, 99.0, 0xFF, 0xFF))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]?.ixf).toBe(-1);
  });

  it('handles iStyleRef == 0 (no style) without sign extension', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealStyled(0, 1.0, 0x00, 0x00))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]?.ixf).toBe(0);
  });
});
