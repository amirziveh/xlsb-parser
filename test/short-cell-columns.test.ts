import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import { buildXlsb, rowHeader, rec, u32, concat } from './helpers';

// §2.2 audit claim: "prevCol = col || 0 corrupts short-cell columns".
// Verification: after a fresh row header (prevCol=-1), col = 0. col || 0 = 0.
// Plain `prevCol = col` would also be 0. The `|| 0` is dead code, not a bug.
// This test LOCKS the existing correct behavior so we can safely remove `|| 0`.
describe('short-cell column sequencing', () => {
  it('places the first short cell of a row at column 0', async () => {
    // BRT_SHORT_RK = 0x0D. data: 2 bytes col+reserved(2 bytes iStyleRef @2..3) + 4 bytes iStyleRef
    // Actually short form: 2 bytes (col? no — short form has implicit col via prevCol+1)
    // Per MS-XLSB BrtCellRkShort: 4 bytes iStyleRef(4)... no wait, short form literally has:
    //   iStyleRef (4 bytes uint32) + value (4 bytes RK)
    // length=8. But code reads: ixf = readU16(d, 2); off=4; reads RK at offset 4.
    // So format tested: col(2 bytes... actually for short form, code reads ixf from offset 2)
    // Looking at code:
    //   const ixf = d.length >= 4 ? readU16(d, 2) : undefined;
    //   const cell = readShortCell(r.type, d, 4, ss);
    // So short record data: bytes 0..1 = something(2), bytes 2..3 = iStyleRef, bytes 4.. = value
    // For BRT_SHORT_RK: value = uint32 RK at offset 4
    // Build a short RK cell encoding 0 as RK → value 0
    const shortRk = (rkValue: number) =>
      rec(0x0d, concat(new Uint8Array([0, 0]), new Uint8Array([0, 0]), u32(rkValue)));

    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [
        concat(
          rowHeader(0),
          shortRk(0), // cell at col 0
          shortRk(0x00000002), // RK encoding integer 0 (same value, but next col 1)
          shortRk(0x00000002), // col 2
        ),
      ],
    });
    const wb = await parseXlsb(xlsb);
    const cols = Object.keys(wb.sheets[0].rows[0].cols)
      .map(Number)
      .sort((a, b) => a - b);
    expect(cols).toEqual([0, 1, 2]);
  });
});
