import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import {
  buildXlsb,
  rowHeader,
  concat,
  cellIsst,
  cellReal,
  sstBytes,
  workbookBinRecordLegacy,
} from './helpers';

// Multi-sheet: covers the per-sheet loop in index.ts and the
// parseWorkbook multi-record walk.
describe('multi-sheet workbooks', () => {
  it('parses 3 sheets in order with distinct names', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['Alpha', 'Beta', 'Gamma'],
      sharedStrings: ['a', 'b', 'c'],
      sheetRecords: [
        concat(rowHeader(0), cellIsst(0, 0)),
        concat(rowHeader(0), cellIsst(0, 1)),
        concat(rowHeader(0), cellIsst(0, 2)),
      ],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets.map((s) => s.name)).toEqual(['Alpha', 'Beta', 'Gamma']);
    expect(wb.sheets[0].rows[0].cols[0]?.v).toBe('a');
    expect(wb.sheets[2].rows[0].cols[0]?.v).toBe('c');
  });

  it('skips a missing sheetN.bin part silently', async () => {
    // buildXlsb always emits every sheetN, so we simulate "missing" by
    // passing an empty body for sheet 2 (the parser still registers it).
    const xlsb = buildXlsb({
      sheetNames: ['Present', 'Empty'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellReal(0, 1)), new Uint8Array(0)],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets.length).toBe(2);
    expect(wb.sheets[1].rows.length).toBe(0);
  });

  it('reports totalCells per sheet', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [
        concat(
          rowHeader(0),
          cellReal(0, 1),
          cellReal(1, 2),
          cellReal(2, 3),
          rowHeader(1),
          cellReal(0, 4),
        ),
      ],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].totalCells).toBe(4);
  });
});

// Legacy BRT_BUNDLE_SH (0x9C) — older encoding that parseWorkbook still
// recognises. Real XLSB files use 0x0E01 but the wire format keeps the legacy
// path for compatibility with old Office versions.
describe('legacy workbook bundle record (BRT_BUNDLE_SH = 0x9C)', () => {
  it('parses sheet names from the legacy form', async () => {
    const sheetNames = ['Legacy1', 'Legacy2'];
    const xlsb = buildXlsb({
      sheetNames,
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellReal(0, 1)), concat(rowHeader(0), cellReal(0, 2))],
      workbookBin: workbookBinRecordLegacy(sheetNames),
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets.map((s) => s.name)).toEqual(['Legacy1', 'Legacy2']);
  });
});

// §2.5 audit claim: "readRichString mishandles rich-text runs". The audit
// said the function only skips the flags byte (1 byte) then reads cch+chars
// at offset+1. Per MS-XLSB §2.5.31 BrtRichStr layout:
//   flags(1) + cch(4) + chars(cch*2) [+ cRun(2) + Run*4 if fRt] [+ phonetic]
// So the chars are read correctly; runs/phonetic come AFTER chars. The audit
// concern was unfounded — current code reads the string correctly, just
// discards formatting runs. These tests verify that.
describe('BrtRichStr with formatting runs (§2.5 audit verification)', () => {
  it('reads the string correctly from a rich-text SST item with fRt=1, cRun=0', async () => {
    // SST item with fRt=1 (has runs) but cRun=0 — exercises the path where
    // runs section is present (just the cRun=0 header). String value must
    // still decode cleanly.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [], // we override the SST below via extraEntries? No —
      // buildXlsb only emits sharedStrings.bin if sharedStrings.length > 0.
      // Workaround: pass a placeholder SST and override via extraEntries.
      sheetRecords: [concat(rowHeader(0), cellIsst(0, 0))],
      extraEntries: {
        'xl/sharedStrings.bin': sstBytes([{ s: 'Rich Text', rich: true }]),
      },
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sharedStrings).toEqual(['Rich Text']);
    expect(wb.sheets[0].rows[0].cols[0]?.v).toBe('Rich Text');
  });

  it('reads a mix of plain and rich SST items in the same file', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [], // overridden below
      sheetRecords: [
        concat(rowHeader(0), cellIsst(0, 0), cellIsst(1, 1), rowHeader(1), cellIsst(0, 2)),
      ],
      extraEntries: {
        'xl/sharedStrings.bin': sstBytes([
          { s: 'plain' },
          { s: 'rich', rich: true },
          { s: 'plain again' },
        ]),
      },
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sharedStrings).toEqual(['plain', 'rich', 'plain again']);
    expect(wb.sheets[0].rows[0].cols[0]?.v).toBe('plain');
    expect(wb.sheets[0].rows[0].cols[1]?.v).toBe('rich');
    expect(wb.sheets[0].rows[1].cols[0]?.v).toBe('plain again');
  });

  it('reads an empty SST (no sharedStrings.bin part at all)', async () => {
    // When there's no SST, ISST cells should still produce a [SST#N] fallback.
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [concat(rowHeader(0), cellIsst(0, 0))],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sharedStrings).toEqual([]);
    expect(wb.sheets[0].rows[0].cols[0]).toMatchObject({ t: 's', v: '[SST#0]' });
  });
});
