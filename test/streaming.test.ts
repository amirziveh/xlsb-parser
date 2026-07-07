import { describe, it, expect } from 'vitest';
import { openXlsb } from '../src/index.js';
import {
  buildXlsb, rowHeader, cellReal, cellIsst, concat,
} from './helpers';

// P4b: streaming via openXlsb() handle. Memory: O(cells_per_row) instead
// of O(total_rows). Tests verify the handle API and that rows are yielded
// in order, complete with their cells attached.
describe('openXlsb — streaming handle', () => {
  it('exposes sheetNames without eagerly parsing sheets', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['Sheet1', 'Sheet2'],
      sharedStrings: [],
      sheetRecords: [new Uint8Array(0), new Uint8Array(0)],
    });
    const handle = await openXlsb(xlsb);
    expect(handle.sheetNames).toEqual(['Sheet1', 'Sheet2']);
    expect(handle.sharedStrings).toEqual([]);
    // buildXlsb always emits a stub styles.bin (rec 0x0000) → styles parses
    // to an empty StylesTable (cellXfs: [], numFmts: empty map).
    expect(handle.styles).not.toBeNull();
    expect(handle.styles?.cellXfs).toEqual([]);
  });

  it('yields rows from iterSheetRows one at a time', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(
        rowHeader(0), cellReal(0, 1.1),
        rowHeader(1), cellReal(0, 2.2),
        rowHeader(2), cellReal(0, 3.3),
      )],
    });
    const handle = await openXlsb(xlsb);
    const rows = [];
    for await (const r of handle.iterSheetRows(0)) rows.push(r);
    expect(rows.length).toBe(3);
    expect(rows[0].row).toBe(0);
    expect(rows[1].row).toBe(1);
    expect(rows[2].row).toBe(2);
    expect(rows[0].cols[0]?.v).toBe(1.1);
    expect(rows[2].cols[0]?.v).toBe(3.3);
  });

  it('yields each row complete with all of its cells attached', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: ['a', 'b'],
      sheetRecords: [concat(
        rowHeader(0),
        cellIsst(0, 0), cellIsst(1, 1), cellReal(2, 100),
        rowHeader(1),
        cellReal(0, 99),
      )],
    });
    const handle = await openXlsb(xlsb);
    const rows = [];
    for await (const r of handle.iterSheetRows(0)) rows.push(r);
    expect(Object.keys(rows[0].cols).length).toBe(3);
    expect(rows[0].cols[0]?.v).toBe('a');
    expect(rows[0].cols[1]?.v).toBe('b');
    expect(rows[0].cols[2]?.v).toBe(100);
    expect(rows[1].cols[0]?.v).toBe(99);
  });

  it('respects maxRows option, stopping early', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(
        rowHeader(0), cellReal(0, 1),
        rowHeader(1), cellReal(0, 2),
        rowHeader(2), cellReal(0, 3),
        rowHeader(3), cellReal(0, 4),
        rowHeader(4), cellReal(0, 5),
      )],
    });
    const handle = await openXlsb(xlsb);
    const rows = [];
    for await (const r of handle.iterSheetRows(0, { maxRows: 3 })) rows.push(r);
    expect(rows.length).toBe(3);
    expect(rows[0].row).toBe(0);
    expect(rows[2].row).toBe(2);
  });

  it('collectSheet drains rows into a Sheet object', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(
        rowHeader(0), cellReal(0, 1), cellReal(1, 2),
        rowHeader(1), cellReal(0, 3),
      )],
    });
    const handle = await openXlsb(xlsb);
    const sheet = await handle.collectSheet(0);
    expect(sheet.name).toBe('S');
    expect(sheet.rows.length).toBe(2);
    expect(sheet.totalCells).toBe(3);
  });

  it('returns an empty iterator when sheetIndex has no part', async () => {
    // buildXlsb always emits sheet parts, so we test sheetIndex out of range.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [new Uint8Array(0)],
    });
    const handle = await openXlsb(xlsb);
    const rows = [];
    for await (const r of handle.iterSheetRows(99)) rows.push(r);
    expect(rows.length).toBe(0);
  });

  it('large sheet: yields 1000 rows without buffering them all in memory', async () => {
    // Build a sheet with 1000 rows.
    const rowRecs: Uint8Array[] = [];
    for (let i = 0; i < 1000; i++) {
      rowRecs.push(concat(rowHeader(i), cellReal(0, i)));
    }
    const xlsb = buildXlsb({
      sheetNames: ['Big'],
      sharedStrings: [],
      sheetRecords: [concat(...rowRecs)],
    });
    const handle = await openXlsb(xlsb);
    let count = 0;
    let first = -1, last = -1;
    for await (const r of handle.iterSheetRows(0)) {
      if (count === 0) first = r.row;
      last = r.row;
      count++;
      if (count % 100 === 0) await new Promise(res => setTimeout(res, 0));
    }
    expect(count).toBe(1000);
    expect(first).toBe(0);
    expect(last).toBe(999);
  });

  it('accepts the legacy 2nd-arg function form for backwards compat', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [], sheetRecords: [new Uint8Array(0)],
    });
    const calls: string[] = [];
    const handle = await openXlsb(xlsb, (msg) => calls.push(msg));
    expect(handle.sheetNames).toEqual(['S']);
    expect(calls.length).toBeGreaterThan(0);
  });

  it('accepts ArrayBuffer input directly', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellReal(0, 1))],
    });
    // Pass the underlying ArrayBuffer (not the Uint8Array view).
    const handle = await openXlsb(xlsb.buffer);
    expect(handle.sheetNames).toEqual(['S']);
  });

  it('throws when the input is not a valid ZIP', async () => {
    await expect(openXlsb(new Uint8Array([1, 2, 3]))).rejects.toThrow('ZIP');
  });

  it('throws when maxZipBytes is exceeded', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellReal(0, 1))],
    });
    const { XlsbSizeError } = await import('../src/index.js');
    await expect(openXlsb(xlsb, { maxZipBytes: 1 })).rejects.toBeInstanceOf(XlsbSizeError);
  });

  it('fires onProgress callback during iteration over a large sheet', async () => {
    const rowRecs: Uint8Array[] = [];
    for (let i = 0; i < 2000; i++) rowRecs.push(concat(rowHeader(i), cellReal(0, i)));
    const xlsb = buildXlsb({
      sheetNames: ['Big'], sharedStrings: [], sheetRecords: [concat(...rowRecs)],
    });
    const handle = await openXlsb(xlsb);
    const calls: { msg: string; pct: number }[] = [];
    for await (const _ of handle.iterSheetRows(0, {
      onProgress: (msg, pct) => calls.push({ msg, pct }),
    })) { /* drain */ }
    expect(calls.length).toBeGreaterThan(0);
  });

  it('collectSheet drains even when onProgress is set', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellReal(0, 1), rowHeader(1), cellReal(0, 2))],
    });
    const handle = await openXlsb(xlsb);
    const calls: { msg: string; pct: number }[] = [];
    const sheet = await handle.collectSheet(0, { onProgress: (msg, pct) => calls.push({ msg, pct }) });
    expect(sheet.rows.length).toBe(2);
  });
});
