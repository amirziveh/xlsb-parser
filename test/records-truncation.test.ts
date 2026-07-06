import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import { buildXlsb, rowHeader, rec, encType, encSize, concat, u32, u16le } from './helpers';

// §2.4: records() silently clamps a record's declared size down to whatever
// bytes remain in the buffer: `if (off + s > data.length) s = data.length - off;`
// A truncated/malformed .bin file therefore produces garbage records with no
// error signal. For a binary parser fed untrusted inputs, this is a security
// and correctness concern — host code should know the file is malformed.
describe('records() rejects truncated .bin files', () => {
  it('throws when a record declares more bytes than remain in the buffer', async () => {
    // Build a workbook.bin record whose declared size (200) exceeds the actual
    // 20 bytes that follow. The pre-fix records() generator would silently
    // yield a 20-byte record, causing parseWorkbook to read inaccurate data
    // (or worse, contiguous-buffer over-reads in nested decoders).
    const declaredSize = 200;
    const fakeRecord = concat(encType(0x0E01), encSize(declaredSize), new Uint8Array(20));
    // Pad workbook with one minimal-but-legible bundle record first to keep
    // parseWorkbook happy after the broken one (we want the throw to come
    // from the broken record itself, not from absent sheets).
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), rec(0x01, concat(u32(0), new Uint8Array(4))))],
      extraEntries: { 'xl/workbook.bin': fakeRecord },
    });
    // Override workbook.bin with the corrupt bytes (buildXlsb already wrote
    // a valid one; we rewrote via extraEntries which is applied AFTER, but to
    // be unambiguous we do it explicitly here)
    // (extraEntries is merged into zipInput last, so it overrides.)

    await expect(parseXlsb(xlsb)).rejects.toThrow(/truncat|declared|record/i);
  });

  it('accepts a well-formed record where declared size matches bytes present', async () => {
    // Sanity: a properly-sized record should NOT trip the new throw.
    const xlsb = buildXlsb({
      sheetNames: ['Sheet1'],
      sharedStrings: ['x'],
      sheetRecords: [concat(rowHeader(0), rec(0x07, concat(u32(0), new Uint8Array(4), u32(0))))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].name).toBe('Sheet1');
  });
});
