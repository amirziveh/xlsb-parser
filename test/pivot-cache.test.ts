import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import {
  buildXlsb, rec, u32, u16le, concat,
} from './helpers';

// Minimal pivot-cache fixtures. The current pivot-cache decoder is heuristic
// (P5 of the roadmap will rewrite it spec-first), so these tests only lock
// in the broad behaviour — field-name extraction and the happy-path through
// the records decoder. They push branch coverage up from 0% without
// asserting on the brittle field-type detection.

// Pivot-cache definition record 0x1B81 (BrtPCDIHFields / PCDFields
// approximated): data layout per parsePivotCache = 20 bytes filler, then
// uint32 nameLen at offset 20, then nameLen*2 UTF-16LE chars at offset 24.
function pivotFieldRecordFixed(name: string): Uint8Array {
  const buf = new Uint8Array(20);
  return rec(0x1B81, concat(buf, u32(name.length), u16le(name)));
}

// Pivot-cache record 0x0021 (BrtPCDIRecord-ish) with one u32 field.
function pivotRecordU32(value: number): Uint8Array {
  return rec(0x0021, u32(value));
}

describe('pivot cache parsing', () => {
  it('extracts field names from a minimal pivot cache definition', async () => {
    const def = concat(pivotFieldRecordFixed('Region'));
    const recs = concat(pivotRecordU32(1), rec(0x2101, new Uint8Array(0)));

    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': def,
        'xl/pivotCache/pivotCacheRecords1.bin': recs,
      },
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.pivotCaches.length).toBe(1);
    expect(wb.pivotCaches[0].name).toBe('PivotCache1');
    expect(wb.pivotCaches[0].fieldNames).toContain('Region');
  });

  it('parses up to the 0x2101 end-of-records marker', async () => {
    const def = pivotFieldRecordFixed('Field');
    const recs = concat(
      pivotRecordU32(100), pivotRecordU32(200),
      rec(0x2101, new Uint8Array(0)),  // end marker
      pivotRecordU32(999),              // should be ignored
    );
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': def,
        'xl/pivotCache/pivotCacheRecords1.bin': recs,
      },
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.pivotCaches[0].rowCount).toBeGreaterThan(0);
  });

  it('skips 0x0021 records before they appear (non-matching types)', async () => {
    // Records part with leading unknown record type, then a real record.
    const def = pivotFieldRecordFixed('X');
    const recs = concat(
      rec(0x9999, u32(0xDEADBEEF)),    // unknown, skipped
      pivotRecordU32(5),                 // valid
      rec(0x2101, new Uint8Array(0)),
    );
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': def,
        'xl/pivotCache/pivotCacheRecords1.bin': recs,
      },
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.pivotCaches[0].rows.length).toBe(1);
  });

  it('parses both cache1 and cache2 when both are present', async () => {
    const def1 = pivotFieldRecordFixed('A');
    const recs1 = concat(pivotRecordU32(1), rec(0x2101, new Uint8Array(0)));
    const def2 = pivotFieldRecordFixed('B');
    const recs2 = concat(pivotRecordU32(2), rec(0x2101, new Uint8Array(0)));

    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': def1,
        'xl/pivotCache/pivotCacheRecords1.bin': recs1,
        'xl/pivotCache/pivotCacheDefinition2.bin': def2,
        'xl/pivotCache/pivotCacheRecords2.bin': recs2,
      },
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.pivotCaches.length).toBe(2);
    expect(wb.pivotCaches[0].name).toBe('PivotCache1');
    expect(wb.pivotCaches[1].name).toBe('PivotCache2');
  });

  it('returns an empty pivotCaches array when no caches are present', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.pivotCaches).toEqual([]);
  });
});
