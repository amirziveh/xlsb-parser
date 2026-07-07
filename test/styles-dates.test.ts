import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import {
  buildXlsb, rowHeader, concat, rec, stylesBinRecord,
  cellRealWithStyle, u32,
} from './helpers';

// P4: styles + date detection. When a cell's iStyleRef points to a cellXfs
// whose numFmtId is a date format, the cell gets `numFmtId`, `isDate: true`,
// and `dateValue` (ISO 8601) — but `v` stays the raw serial number.
describe('styles + date metadata', () => {
  it('exposes numFmtId on cells when styles.bin is present', async () => {
    // Single cellXfs with numFmtId = 0 (General, not a date)
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(
        rowHeader(0),
        // cell at col 0, iStyleRef = 0 (first cellXfs entry), REAL = 42.0
        cellRealWithStyle(0, 42.0, 0, 0),
      )],
      extraEntries: {
        // Override the default empty styles.bin
        'xl/styles.bin': stylesBinRecord({ cellXfs: [0] }),
      },
    });
    const wb = await parseXlsb(xlsb);
    const cell = wb.sheets[0].rows[0].cols[0];
    expect(cell?.numFmtId).toBe(0);
    expect(cell?.isDate).toBeUndefined(); // numFmtId=0 is General, not a date
    expect(cell?.dateValue).toBeUndefined();
    expect(cell?.v).toBe(42.0); // raw value unchanged
  });

  it('marks cells with built-in date numFmtId (e.g., 14 = mm-dd-yyyy)', async () => {
    // Build a sheet where serial 44927 (= 2023-01-01) is styled with
    // iStyleRef = 0 → cellXfs[0] = numFmtId 14 (built-in date).
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(
        rowHeader(0),
        cellRealWithStyle(0, 44927, 0, 0),
      )],
      extraEntries: {
        'xl/styles.bin': stylesBinRecord({ cellXfs: [14] }),
      },
    });
    const wb = await parseXlsb(xlsb);
    const cell = wb.sheets[0].rows[0].cols[0];
    expect(cell?.numFmtId).toBe(14);
    expect(cell?.isDate).toBe(true);
    expect(cell?.v).toBe(44927); // serial stays as the number
    expect(cell?.dateValue).toBe('2023-01-01T00:00:00.000Z');
  });

  it('marks cells with built-in time numFmtId 22 (mm/dd/yyyy hh:mm)', async () => {
    // Serial 44927.5 → 2023-01-01T12:00:00Z
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealWithStyle(0, 44927.5, 0, 0))],
      extraEntries: { 'xl/styles.bin': stylesBinRecord({ cellXfs: [22] }) },
    });
    const wb = await parseXlsb(xlsb);
    const cell = wb.sheets[0].rows[0].cols[0];
    expect(cell?.isDate).toBe(true);
    expect(cell?.dateValue).toBe('2023-01-01T12:00:00.000Z');
  });

  it('detects custom date format strings (e.g., "yyyy-mm-dd")', async () => {
    // Custom numFmt strings: register via BrtFmt with a high numFmtId (164+),
    // then create a cellXfs referencing that custom ID.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealWithStyle(0, 44927, 0, 0))],
      extraEntries: {
        'xl/styles.bin': stylesBinRecord({
          customFmts: { 164: 'yyyy-mm-dd' },
          cellXfs: [164],
        }),
      },
    });
    const wb = await parseXlsb(xlsb);
    const cell = wb.sheets[0].rows[0].cols[0];
    expect(cell?.numFmtId).toBe(164);
    expect(cell?.isDate).toBe(true);
    expect(cell?.dateValue).toBe('2023-01-01T00:00:00.000Z');
  });

  it('rejects purely numeric custom formats (e.g., "#,##0.00")', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealWithStyle(0, 1234.5, 0, 0))],
      extraEntries: {
        'xl/styles.bin': stylesBinRecord({
          customFmts: { 165: '#,##0.00' },
          cellXfs: [165],
        }),
      },
    });
    const wb = await parseXlsb(xlsb);
    const cell = wb.sheets[0].rows[0].cols[0];
    expect(cell?.numFmtId).toBe(165);
    expect(cell?.isDate).toBeUndefined();
    expect(cell?.dateValue).toBeUndefined();
    expect(cell?.v).toBe(1234.5);
  });

  it('ignores date tokens inside escaped/quoted sections', async () => {
    // "d" is inside literal quotes → NOT a date
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealWithStyle(0, 5, 0, 0))],
      extraEntries: {
        'xl/styles.bin': stylesBinRecord({
          customFmts: { 166: '"day"d' }, // 'day' is literal, 'd' is outside it → DATE
          cellXfs: [166],
        }),
      },
    });
    const wb = await parseXlsb(xlsb);
    const cell = wb.sheets[0].rows[0].cols[0];
    expect(cell?.isDate).toBe(true); // 'd' outside quotes → date

    // And the inverse: only `\d` escapes → NOT a date
    const xlsb2 = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealWithStyle(0, 5, 0, 0))],
      extraEntries: {
        'xl/styles.bin': stylesBinRecord({
          customFmts: { 167: '\\d #0' }, // `\d` is an escaped 'd' → NOT a date
          cellXfs: [167],
        }),
      },
    });
    const wb2 = await parseXlsb(xlsb2);
    const cell2 = wb2.sheets[0].rows[0].cols[0];
    expect(cell2?.isDate).toBeUndefined();
  });

  it('falls back gracefully when iStyleRef is out of range', async () => {
    // iStyleRef = 99 but styles table only has 1 entry → cell gets no numFmtId
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealWithStyle(0, 5, 99, 0))],
      extraEntries: { 'xl/styles.bin': stylesBinRecord({ cellXfs: [14] }) },
    });
    const wb = await parseXlsb(xlsb);
    const cell = wb.sheets[0].rows[0].cols[0];
    expect(cell?.numFmtId).toBeUndefined();
    expect(cell?.isDate).toBeUndefined();
    expect(cell?.v).toBe(5); // value still decoded
  });

  it('works without any styles.bin present (no metadata added to cells)', async () => {
    // buildXlsb always emits a stub styles.bin — to test "no styles" we need
    // to remove it. Use a custom ZIP without styles.bin.
    const { zipSync } = await import('fflate');
    const text = (s: string) => new TextEncoder().encode(s);
    // Reuse the minimal fixture, then rebuild the zip omitting styles.bin
    const minimal = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealWithStyle(0, 5, 0, 0))],
    });
    const { unzipSync } = await import('fflate');
    const unzipped = unzipSync(minimal);
    delete unzipped['xl/styles.bin'];
    // Also need to delete the styles rel in workbook.bin.rels, but for the
    // purposes of this test we only care that styles is null and the cell
    // still parses without metadata.
    const rebuilt = new Uint8Array(zipSync(unzipped, { level: 0 }));
    const wb = await parseXlsb(rebuilt);
    const cell = wb.sheets[0].rows[0].cols[0];
    expect(cell?.v).toBe(5);
    expect(cell?.numFmtId).toBeUndefined();
    expect(cell?.isDate).toBeUndefined();
  });
});

describe('XlsbSizeError caps', () => {
  it('throws XlsbSizeError when maxZipBytes is exceeded', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRealWithStyle(0, 1, 0, 0))],
    });
    const { XlsbSizeError } = await import('../src/index.js');
    await expect(parseXlsb(xlsb, { maxZipBytes: 1 })).rejects.toBeInstanceOf(XlsbSizeError);
  });

  it('throws XlsbSizeError when maxRowsPerSheet is exceeded', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(
        rowHeader(0), cellRealWithStyle(0, 1, 0, 0),
        rowHeader(1), cellRealWithStyle(0, 2, 0, 0),
        rowHeader(2), cellRealWithStyle(0, 3, 0, 0),
      )],
    });
    const { XlsbSizeError } = await import('../src/index.js');
    await expect(parseXlsb(xlsb, { maxRowsPerSheet: 2 })).rejects.toBeInstanceOf(XlsbSizeError);
  });
});
