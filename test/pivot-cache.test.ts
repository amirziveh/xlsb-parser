import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import {
  buildXlsb, rec, u32, u16le, concat, pcdFieldFull, pcdStr, pcdDate, pcdNum,
  pcdRun, pcdErr, pcRecord, f64,
  pcRecordsHeader, pcRecordsEnd,
} from './helpers';
import { BRT_PCDI_STRING, BRT_PC_RECORD_DT, BRT_PCDINUMBER } from '../src/record-stream.js';

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
  return rec(0x1b81, concat(buf, u32(name.length), u16le(name)));
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
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches.length).toBe(1);
    expect(wb.pivotCaches[0].name).toBe('PivotCache1');
    expect(wb.pivotCaches[0].fieldNames).toContain('Region');
  });

  it('parses up to the 0x2101 end-of-records marker', async () => {
    const def = concat(pcdFieldFull('Field', { isSrc: true, fText: true }));
    const recs = concat(
      pcRecordsHeader(2),
      pcRecord([u32(0)]),
      pcRecord([u32(0)]),
      pcRecordsEnd(),
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
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches[0].rowCount).toBeGreaterThan(0);
  });

  it('skips 0x0021 records before they appear (non-matching types)', async () => {
    // Records part with leading unknown record type, then a real record.
    const def = pivotFieldRecordFixed('X');
    const recs = concat(
      rec(0x9999, u32(0xdeadbeef)), // unknown, skipped
      pivotRecordU32(5), // valid
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
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
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
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches.length).toBe(2);
    expect(wb.pivotCaches[0].name).toBe('PivotCache1');
    expect(wb.pivotCaches[1].name).toBe('PivotCache2');
  });

  it('parses 5+ pivot caches (not hardcoded to 2 — regression for audit §2.7)', async () => {
    const extras: Record<string, Uint8Array> = {};
    for (let i = 1; i <= 5; i++) {
      extras[`xl/pivotCache/pivotCacheDefinition${i}.bin`] = pivotFieldRecordFixed('F' + i);
      extras[`xl/pivotCache/pivotCacheRecords${i}.bin`] = concat(
        pivotRecordU32(i),
        rec(0x2101, new Uint8Array(0)),
      );
    }
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: extras,
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches.length).toBe(5);
    expect(wb.pivotCaches.map(p => p.name)).toEqual([
      'PivotCache1', 'PivotCache2', 'PivotCache3', 'PivotCache4', 'PivotCache5',
    ]);
  });

  it('parses non-contiguous cache numbers (1 and 5, but not 2/3/4)', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': pivotFieldRecordFixed('A'),
        'xl/pivotCache/pivotCacheRecords1.bin': concat(pivotRecordU32(1), rec(0x2101, new Uint8Array(0))),
        'xl/pivotCache/pivotCacheDefinition5.bin': pivotFieldRecordFixed('E'),
        'xl/pivotCache/pivotCacheRecords5.bin': concat(pivotRecordU32(5), rec(0x2101, new Uint8Array(0))),
      },
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches.length).toBe(2);
    expect(wb.pivotCaches[0].name).toBe('PivotCache1');
    expect(wb.pivotCaches[1].name).toBe('PivotCache5');
  });

  it('skips a pivotCacheDefinition without a matching records part', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': pivotFieldRecordFixed('Orphan'),
        // intentionally no pivotCacheRecords1.bin
      },
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches).toEqual([]);
  });

  it('returns an empty pivotCaches array when no caches are present', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.pivotCaches).toEqual([]);
    // Also true even when opt-in is set: no caches in the file.
    const wb2 = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb2.pivotCaches).toEqual([]);
  });
});

describe('pivot cache field descriptors', () => {
  function buildPc(name: string, def: Uint8Array, recs: Uint8Array) {
    return buildXlsb({
      sheetNames: ['S'], sharedStrings: [], sheetRecords: [],
      extraEntries: {
        [`xl/pivotCache/pivotCacheDefinition${name}.bin`]: def,
        [`xl/pivotCache/pivotCacheRecords${name}.bin`]: recs,
      },
    });
  }

  it('reads fNumField/fDateInField/fHasTextItem into field.kind', async () => {
    const def = concat(
      pcdFieldFull('Region', { isSrc: true, fText: true }),
      pcdStr('North'), pcdStr('South'),
      pcdFieldFull('Amount', { isSrc: true, fNum: true }),
      pcdFieldFull('When', { isSrc: true, fDate: true }),
      pcdDate(2024, 5, 10, 13, 30, 0),
    );
    const recs = concat(pcRecordsHeader(0), pcRecordsEnd());
    const wb = await parseXlsb(buildPc('1', def, recs), { parsePivotCaches: true });
    const pc = wb.pivotCaches[0];
    expect(pc.fields.map(f => f.kind)).toEqual(['indexed', 'number', 'date']);
    expect(pc.fields[0].sharedItems.map(c => (c as { v: unknown })?.v)).toEqual(['North', 'South']);
    expect(pc.fields[0].isSrc).toBe(true);
  });
});

describe('pivot cache row decoding', () => {
  it('decodes PCDIDateTime with non-zero hour (F4)', async () => {
    const def = concat(
      pcdFieldFull('When', { isSrc: true, fDate: true }),
      pcdDate(2024, 5, 10, 13, 30, 0),
    );
    const rgb = concat(u32(0));
    const recs = concat(pcRecordsHeader(1), pcRecord([rgb]), pcRecordsEnd());
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [], sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': def,
        'xl/pivotCache/pivotCacheRecords1.bin': recs,
      },
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches[0].rows[0][0]).toEqual({ t: 'd', v: '2024-05-10T13:30:00' });
  });

  it('decodes BrtBeginPCDIRun number and error runs (F5)', async () => {
    const def = concat(
      pcdFieldFull('Num', { isSrc: true, fNum: true }),
      pcdRun(0x01, [1.5, 2.5]),
      pcdFieldFull('Err', { isSrc: true, fText: true }),
      pcdRun(0x10, [0x17]), // #REF!
    );
    const rgb = concat(u32(0), u32(0));
    const recs = concat(pcRecordsHeader(1), pcRecord([rgb]), pcRecordsEnd());
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [], sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': def,
        'xl/pivotCache/pivotCacheRecords1.bin': recs,
      },
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches[0].fields[0].sharedItems[1]).toEqual({ t: 'n', v: 2.5 });
    expect(wb.pivotCaches[0].fields[1].sharedItems[0]).toEqual({ t: 'e', v: '#REF!' });
  });

  it('preserves numeric-as-text and non-Latin strings (F8)', async () => {
    const def = concat(
      pcdFieldFull('Code', { isSrc: true, fText: true }),
      pcdStr('12345'),
      pcdStr('مرحبا'),
    );
    const recs = concat(
      pcRecordsHeader(2),
      pcRecord([u32(0)]),
      pcRecord([u32(1)]),
      pcRecordsEnd(),
    );
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [], sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': def,
        'xl/pivotCache/pivotCacheRecords1.bin': recs,
      },
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches[0].rows[0][0]).toEqual({ t: 's', v: '12345' });
    expect(wb.pivotCaches[0].rows[1][0]).toEqual({ t: 's', v: 'مرحبا' });
  });
});

describe('pivot cache PCDIDT mode', () => {
  it('decodes BrtPCRRecordDt rows via per-field BrtPCDI* records (F2)', async () => {
    const def = concat(
      pcdFieldFull('Name', { isSrc: true, fText: true }),
      pcdStr('Alice'),
      pcdFieldFull('Val', { isSrc: true, fNum: true }),
    );
    const row = concat(
      rec(BRT_PC_RECORD_DT, new Uint8Array(0)),
      rec(BRT_PCDI_STRING, concat(u32(5), u16le('Alice'))),
      rec(BRT_PCDINUMBER, f64(3.14)),
    );
    const recs = concat(pcRecordsHeader(1), row, pcRecordsEnd());
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [], sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': def,
        'xl/pivotCache/pivotCacheRecords1.bin': recs,
      },
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches[0].rows.length).toBe(1);
    expect(wb.pivotCaches[0].rows[0]).toEqual([{ t: 's', v: 'Alice' }, { t: 'n', v: 3.14 }]);
  });
});
