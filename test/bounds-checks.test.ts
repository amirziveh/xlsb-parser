import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import { buildXlsb, rowHeader, rec, u32, concat } from './helpers';

// §2.3: readCell/readShortCell called readU32/readF64/d[off] without bounds
// checks, so a malformed (too-short) cell record would throw RangeError out
// of parseXlsb (via DataView constructor). Real-world cause: truncated files,
// fuzzed inputs, or partial-write recovery. Expected behavior: skip the bad
// cell, return everything else. The whole sheet should not blow up.
describe('bounds checks on truncated cell records', () => {
  it('does NOT throw on a truncated BRT_CELL_REAL record (only 4 bytes data)', async () => {
    // BrtCellReal normally: col(4) + iStyleRef+reserved(4) + f64(8) = 16 bytes.
    // Truncated to 4.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), rec(0x05, u32(0)))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]).toBeUndefined();
  });

  it('does NOT throw on a truncated BRT_CELL_ISST record (no SST index)', async () => {
    // BrtCellIsst: col(4) + iStyleRef+reserved(4) + isst(4) = 12 bytes. Truncated to 4.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), rec(0x07, u32(0)))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]).toBeUndefined();
  });

  it('does NOT throw on a truncated BRT_CELL_BOOL record (missing bool byte)', async () => {
    // BrtCellBool: col(4) + iStyleRef+reserved(4) + bool(1) = 9 bytes. Truncated to 4.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), rec(0x04, u32(0)))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]).toBeUndefined();
  });

  it('still parses valid cells surrounding a truncated one', async () => {
    // Truncated REAL between two valid ISST cells. Only the bad cell is dropped.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: ['A'],
      sheetRecords: [
        concat(
          rowHeader(0),
          rec(0x07, concat(u32(0), new Uint8Array(4), u32(0))), // valid isst → "A" at col 0
          rec(0x05, u32(2)), // truncated real at col 2
          rec(0x07, concat(u32(3), new Uint8Array(4), u32(0))), // valid isst → "A" at col 3
        ),
      ],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]?.v).toBe('A');
    expect(wb.sheets[0].rows[0].cols[2]).toBeUndefined();
    expect(wb.sheets[0].rows[0].cols[3]?.v).toBe('A');
  });
});
