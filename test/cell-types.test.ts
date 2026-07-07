import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import {
  buildXlsb, rowHeader, concat,
  cellRk, cellError, cellBool, cellStringInline, cellFmlaNum, cellFmlaString, cellFmlaBool, cellFmlaError,
  shortBlank, shortRk, shortError, shortBool, shortReal, shortString, shortIsst,
  u32,
} from './helpers';

// Comprehensive coverage of every cell record type per MS-XLSB §2.4.
// Each long-form type is exercised at col 0 with a representative value; each
// short-form type at col 1 (since short form uses prevCol+1 and prevCol=0
// from the long cell immediately preceding).

describe('long-form cell types', () => {
  it('BRT_CELL_RK decodes integer RK to a number', async () => {
    // fInt=1, num=100, fx100=0 → integer 100
    // bits: 0b100 | (1<<1) = 0b110 = 6
    // RK encoding: bit0=fx100, bit1=fInt, bits2..31=num. 100<<2 | 0b10 = 402 = 0x192
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellRk(0, (100 << 2) | 0b10))],
    });
    expect((await parseXlsb(xlsb)).sheets[0].rows[0].cols[0]).toMatchObject({ t: 'n', v: 100 });
  });

  it('BRT_CELL_BOOL decodes TRUE/FALSE', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellBool(0, true), cellBool(1, false))],
    });
    const r = (await parseXlsb(xlsb)).sheets[0].rows[0];
    expect(r.cols[0]).toMatchObject({ t: 'b', v: true });
    expect(r.cols[1]).toMatchObject({ t: 'b', v: false });
  });

  it('BRT_CELL_ERROR decodes every Excel error code', async () => {
    // 0x00 #NULL!  0x07 #DIV/0!  0x0F #VALUE!  0x17 #REF!  0x1D #NAME?
    // 0x24 #NUM!  0x2A #N/A  0x2B #GETTING_DATA
    const codes = [0x00, 0x07, 0x0F, 0x17, 0x1D, 0x24, 0x2A, 0x2B];
    const expected = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'];
    const cells = codes.map((c, i) => cellError(i, c));
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), ...cells)],
    });
    const row = (await parseXlsb(xlsb)).sheets[0].rows[0];
    codes.forEach((_, i) => {
      expect(row.cols[i]).toMatchObject({ t: 'e', err: expected[i] });
    });
  });

  it('BRT_CELL_ERROR with unknown code produces #ERR(n)', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellError(0, 0x99))],
    });
    const cell = (await parseXlsb(xlsb)).sheets[0].rows[0].cols[0];
    expect(cell?.t).toBe('e');
    expect(cell?.err).toBe('#ERR(153)');
  });

  it('BRT_CELL_ST decodes an inline rich string', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellStringInline(0, 'inline!'))],
    });
    expect((await parseXlsb(xlsb)).sheets[0].rows[0].cols[0]).toMatchObject({ t: 's', v: 'inline!' });
  });

  it('BRT_CELL_ISST exists in the long-form family (covered by minimal fixture)', async () => {
    // Sanity check: a single ISST cell resolves from shared strings
    const { cellIsst } = await import('./helpers');
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: ['hello'],
      sheetRecords: [concat(rowHeader(0), cellIsst(0, 0))],
    });
    expect((await parseXlsb(xlsb)).sheets[0].rows[0].cols[0]).toMatchObject({ t: 's', v: 'hello' });
  });

  it('BRT_CELL_ISST with out-of-range index produces [SST#N]', async () => {
    // Reuse cellIsst-like builder but with an SST index that doesn't exist
    const { cellIsst } = await import('./helpers');
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: ['only-one'],
      sheetRecords: [concat(rowHeader(0), cellIsst(0, 999))],
    });
    expect((await parseXlsb(xlsb)).sheets[0].rows[0].cols[0]).toMatchObject({ t: 's', v: '[SST#999]' });
  });
});

describe('formula cell types', () => {
  it('BRT_FMLA_NUM decodes a float64', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellFmlaNum(0, 3.14))],
    });
    expect((await parseXlsb(xlsb)).sheets[0].rows[0].cols[0]).toMatchObject({ t: 'n', v: 3.14 });
  });

  it('BRT_FMLA_STRING decodes a wide string', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellFmlaString(0, '=SUM(A1)'))],
    });
    expect((await parseXlsb(xlsb)).sheets[0].rows[0].cols[0]).toMatchObject({ t: 's', v: '=SUM(A1)' });
  });

  it('BRT_FMLA_BOOL decodes TRUE', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellFmlaBool(0, true))],
    });
    expect((await parseXlsb(xlsb)).sheets[0].rows[0].cols[0]).toMatchObject({ t: 'b', v: true });
  });

  it('BRT_FMLA_ERROR decodes #DIV/0!', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellFmlaError(0, 0x07))],
    });
    expect((await parseXlsb(xlsb)).sheets[0].rows[0].cols[0]).toMatchObject({ t: 'e', err: '#DIV/0!' });
  });
});

describe('short-form cell types', () => {
  it('decodes a row of every SHORT_* variant in sequence', async () => {
    // Long cell at col 0 sets prevCol=0; short cells then go to cols 1,2,3,...
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: ['short-string'],
      // Need to start each short category from a long cell to reset prevCol=0
      sheetRecords: [concat(
        rowHeader(0),
        // Set 1: short form starting after a long cell at col 0
        cellStringInline(0, 'X'),     // long cell at col 0, prevCol=0
        shortBlank(),                 // col 1
        // Start a new row to reset prevCol=-1, then a long cell at col 0
        // but wait — short form uses prevCol+1. After a long cell at col 0,
        // short cell goes to col 1. Let's just test each type explicitly.

        rowHeader(1),
        cellStringInline(0, 'X'),     // long at col 0 (prevCol=0)
        shortIsst(0),                 // short at col 1 → SST[0] = 'short-string'

        rowHeader(2),
        cellStringInline(0, 'X'),
        shortRk((50 << 2) | 0b10),    // 50 as integer RK, col 1

        rowHeader(3),
        cellStringInline(0, 'X'),
        shortError(0x07),             // #DIV/0! at col 1

        rowHeader(4),
        cellStringInline(0, 'X'),
        shortBool(true),              // TRUE at col 1

        rowHeader(5),
        cellStringInline(0, 'X'),
        shortReal(99.5),              // f64 at col 1

        rowHeader(6),
        cellStringInline(0, 'X'),
        shortString('short!'),        // inline string at col 1
      )],
    });
    const rows = (await parseXlsb(xlsb)).sheets[0].rows;
    expect(rows[1].cols[1]).toMatchObject({ t: 's', v: 'short-string' });
    expect(rows[2].cols[1]).toMatchObject({ t: 'n', v: 50 });
    expect(rows[3].cols[1]).toMatchObject({ t: 'e', err: '#DIV/0!' });
    expect(rows[4].cols[1]).toMatchObject({ t: 'b', v: true });
    expect(rows[5].cols[1]).toMatchObject({ t: 'n', v: 99.5 });
    expect(rows[6].cols[1]).toMatchObject({ t: 's', v: 'short!' });
  });
});
