import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import { makeMinimalXlsb } from './helpers';

describe('parseXlsb', () => {
  it('rejects empty buffer with ZIP error', async () => {
    await expect(parseXlsb(new ArrayBuffer(0))).rejects.toThrow('ZIP');
  });

  it('rejects random data as invalid ZIP', async () => {
    await expect(parseXlsb(new Uint8Array([1, 2, 3]))).rejects.toThrow('ZIP');
  });

  it('accepts ArrayBuffer input', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb.buffer);
    expect(wb.sheets[0].name).toBe('Sheet1');
  });

  it('accepts Uint8Array input', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].name).toBe('Sheet1');
  });
});

describe('minimal XLSB fixture', () => {
  it('generates a non-trivial ZIP', () => {
    const xlsb = makeMinimalXlsb();
    expect(xlsb.length).toBeGreaterThan(500);
  });
});

describe('sheet parsing', () => {
  it('finds one sheet with correct name', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets.length).toBe(1);
    expect(wb.sheets[0].name).toBe('Sheet1');
  });

  it('parses 3 rows', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows.length).toBe(3);
  });

  it('assigns correct row indices', async () => {
    const xlsb = makeMinimalXlsb();
    const rows = await parseXlsb(xlsb).then((w) => w.sheets[0].rows);
    expect(rows[0].row).toBe(0);
    expect(rows[1].row).toBe(1);
    expect(rows[2].row).toBe(2);
  });
});

describe('shared strings', () => {
  it('extracts all shared strings', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    expect(wb.sharedStrings).toEqual(['Name', 'Value', 'foo', 'bar']);
  });

  it('resolves shared string cell values', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    const rows = wb.sheets[0].rows;
    expect(rows[0].cols[0]?.v).toBe('Name');
    expect(rows[0].cols[0]?.t).toBe('s');
    expect(rows[1].cols[0]?.v).toBe('foo');
    expect(rows[2].cols[0]?.v).toBe('bar');
  });
});

describe('cell value types', () => {
  it('reads a real (float64) cell', async () => {
    const xlsb = makeMinimalXlsb();
    const rows = await parseXlsb(xlsb).then((w) => w.sheets[0].rows);
    expect(rows[1].cols[1]?.v).toBe(42.5);
    expect(rows[1].cols[1]?.t).toBe('n');
  });

  it('reads a blank cell', async () => {
    const xlsb = makeMinimalXlsb();
    const rows = await parseXlsb(xlsb).then((w) => w.sheets[0].rows);
    expect(rows[2].cols[1]?.t).toBe('blank');
    expect(rows[2].cols[1]?.v).toBeUndefined();
  });
});

describe('progress callback', () => {
  it('calls callback with progress updates', async () => {
    const xlsb = makeMinimalXlsb();
    const calls: { msg: string; pct: number }[] = [];
    await parseXlsb(xlsb, (msg, pct) => calls.push({ msg, pct }));
    expect(calls.length).toBeGreaterThan(0);
    expect(calls[calls.length - 1].msg).toBe('Done');
    expect(calls[calls.length - 1].pct).toBe(100);
  });
});

describe('summary', () => {
  it('reports file and record counts when dumpBinaries is opt-in', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb, { dumpBinaries: true });
    expect(wb.summary.fileCount).toBeGreaterThan(0);
    expect(wb.summary.totalRecords).toBeGreaterThan(0);
  });

  it('counts .bin files in summary.fileCount even without dumpBinaries', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    expect(wb.summary.fileCount).toBeGreaterThan(0);
    // totalRecords is 0 because dumpBinaries is now opt-in (P4)
    expect(wb.summary.totalRecords).toBe(0);
  });
});
