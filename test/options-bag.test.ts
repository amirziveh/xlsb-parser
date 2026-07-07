import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import { makeMinimalXlsb } from './helpers';

// P4a: parseXlsb signature change.
// New signature: parseXlsb(data, options?) where options is an object.
// Backwards-compat: if 2nd arg is a function, treat it as onProgress (1.x).
// The legacy form will be removed at 2.0.
describe('parseXlsb signature (P4a options bag)', () => {
  it('calls onProgress passed via options.onProgress', async () => {
    const xlsb = makeMinimalXlsb();
    const calls: { msg: string; pct: number }[] = [];
    await parseXlsb(xlsb, { onProgress: (msg, pct) => calls.push({ msg, pct }) });
    expect(calls.length).toBeGreaterThan(0);
    expect(calls[calls.length - 1].msg).toBe('Done');
    expect(calls[calls.length - 1].pct).toBe(100);
  });

  it('accepts the legacy form (2nd arg = function) for backwards compat', async () => {
    const xlsb = makeMinimalXlsb();
    const calls: { msg: string; pct: number }[] = [];
    // Legacy form — 2nd arg is the callback directly. TS won't allow this
    // post-typing change since the signature is now options-only, so cast.
    await (
      parseXlsb as unknown as (
        d: Uint8Array,
        cb: (msg: string, pct: number) => void,
      ) => Promise<unknown>
    )(xlsb, (msg, pct) => calls.push({ msg, pct }));
    expect(calls.length).toBeGreaterThan(0);
    expect(calls[calls.length - 1].msg).toBe('Done');
  });

  it('accepts no options arg at all (default behavior)', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    expect(wb.sheets[0].name).toBe('Sheet1');
  });

  it('populates binaryDumps by default? No — P4 makes it opt-in', async () => {
    // P4 default: dumpBinaries is false. binaryDumps is [].
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    expect(wb.binaryDumps).toEqual([]);
  });

  it('populates binaryDumps when dumpBinaries: true', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb, { dumpBinaries: true });
    expect(wb.binaryDumps.length).toBeGreaterThan(0);
  });

  it('populates xmlFiles by default? No — P4 makes it opt-in', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb);
    expect(wb.xmlFiles).toEqual({});
  });

  it('populates xmlFiles when readXml: true', async () => {
    const xlsb = makeMinimalXlsb();
    const wb = await parseXlsb(xlsb, { readXml: true });
    expect(Object.keys(wb.xmlFiles).length).toBeGreaterThan(0);
    expect(wb.xmlFiles['[Content_Types].xml']).toContain('<Types');
  });

  it('skips pivot caches by default (parsePivotCaches: false)', async () => {
    // Even if pivotCache parts were present, default skips them.
    const { buildXlsb, rec, u32, concat } = await import('./helpers');
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': rec(
          0x1b81,
          concat(new Uint8Array(20), u32(1), new Uint8Array([0, 0])),
        ),
        'xl/pivotCache/pivotCacheRecords1.bin': rec(0x2101, new Uint8Array(0)),
      },
    });
    const wb = await parseXlsb(xlsb);
    expect(wb.pivotCaches).toEqual([]);
  });

  it('parses pivot caches when parsePivotCaches: true', async () => {
    const { buildXlsb, rec, u32, concat } = await import('./helpers');
    const xlsb = buildXlsb({
      sheetNames: ['S'],
      sharedStrings: [],
      sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': concat(
          rec(0x1b81, concat(new Uint8Array(20), u32(1), new Uint8Array([0, 0]))),
        ),
        'xl/pivotCache/pivotCacheRecords1.bin': rec(0x2101, new Uint8Array(0)),
      },
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    expect(wb.pivotCaches.length).toBe(1);
    expect(wb.pivotCaches[0].name).toBe('PivotCache1');
  });
});
