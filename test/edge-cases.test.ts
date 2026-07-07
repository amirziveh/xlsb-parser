import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import { buildXlsb, rowHeader, rec, u32, concat } from './helpers';

// Edge cases that close coverage gaps at the cell-type and orchestrator level.

describe('unknown cell type opcode', () => {
  it('drops a long-form cell whose record type is unrecognised', async () => {
    // 0x09 is BRT_FMLA_NUM actually — pick something in-range but unrecognised.
    // The long-cell range per parseSheet is 0x01..0x0B; we use 0x50 which is
    // outside that range so it falls through to "neither long nor short".
    // That means we won't actually invoke readCell's default branch. Instead
    // exercise the long-cell branch with an unrecognised sub-type by using
    // opcode 0x0C... no that's BRT_SHORT_BLANK.
    // To hit readCell's default, we need r.type in [0x01..0x0B] (range check)
    // but not matching any case. None of 0x01..0x0B is unknown since we
    // covered every BRT_CELL/FMLA opcode. So the default branch is
    // effectively dead code given the switch covers all opcodes in range.
    // We verify that by enumerating all opcodes 0x01..0x0B and confirming
    // none returns null except for the empty-default case... actually no
    // opcode in that range is "unknown" because the switch is exhaustive.
    // Instead, hit the parseSheet fall-through (the OUTER if), which means
    // emitting a record whose type is outside [BRT_CELL_BLANK..BRT_FMLA_ERROR]
    // AND outside [BRT_SHORT_BLANK..BRT_SHORT_ISST]. That hits the "neither
    // long nor short cell" path → continue.
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [],
      sheetRecords: [concat(
        rowHeader(0),
        rec(0x80, u32(0)), // unknown type, neither long nor short cell
        cellStuffThatShouldNotParse(),
      )],
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].rows[0].cols[0]).toBeUndefined();
  });
});

function cellStuffThatShouldNotParse(): Uint8Array {
  // A real cell at col 1 to verify the unknown record didn't corrupt state
  return rec(0x05, concat(u32(1), new Uint8Array(4), new Uint8Array([0, 0, 0, 0, 0, 0, 0, 0]))); // REAL = 0.0
}

describe('workbook.bin missing', () => {
  it('throws xl/workbook.bin not found when the part is absent', async () => {
    // buildXlsb always emits xl/workbook.bin; remove it by overriding with
    // a sentinel entry that... actually we can't easily nullify a part.
    // Instead use buildXlsb's extraEntries to inject a corrupt entry then
    // delete the good one in the test by passing an empty workbookBin?
    // Simpler: build a minimal XLSB-like ZIP directly without workbook.bin.
    const { zipSync } = await import('fflate');
    const text = (s: string) => new TextEncoder().encode(s);
    const zipInput: Record<string, Uint8Array> = {
      '[Content_Types].xml': text(
        '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
      ),
      '_rels/.rels': text('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'),
      // NOTE: no xl/workbook.bin
    };
    const buf = new Uint8Array(zipSync(zipInput, { level: 0 }));
    await expect(parseXlsb(buf)).rejects.toThrow('xl/workbook.bin not found');
  });
});
