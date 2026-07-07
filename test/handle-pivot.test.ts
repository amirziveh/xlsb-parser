import { describe, it, expect } from 'vitest';
import { openXlsb } from '../src/index.js';
import {
  buildXlsb,
  pcdFieldFull,
  pcdStr,
  pcRecordsHeader,
  pcRecord,
  pcRecordsEnd,
  concat,
  u32,
} from './helpers';

function pcXlsb(): Uint8Array {
  const def = concat(
    pcdFieldFull('Region', { isSrc: true, fText: true }),
    pcdStr('North'),
    pcdStr('South'),
  );
  const recs = concat(
    pcRecordsHeader(3),
    pcRecord([u32(0)]),
    pcRecord([u32(1)]),
    pcRecord([u32(0)]),
    pcRecordsEnd(),
  );
  return buildXlsb({
    sheetNames: ['S'],
    sharedStrings: [],
    sheetRecords: [],
    extraEntries: {
      'xl/pivotCache/pivotCacheDefinition1.bin': def,
      'xl/pivotCache/pivotCacheRecords1.bin': recs,
    },
  });
}

describe('openXlsb pivot streaming', () => {
  it('exposes eager defs and streams rows via iterPivotCacheRows', async () => {
    const h = await openXlsb(pcXlsb(), { parsePivotCaches: true });
    expect(h.pivotCaches.length).toBe(1);
    expect(h.pivotCaches[0].fieldNames).toEqual(['Region']);
    expect(h.pivotCaches[0].rowCount).toBe(3);
    const got: Array<Record<string, unknown>> = [];
    for await (const row of h.iterPivotCacheRows(0)) {
      got.push(row as unknown as Record<string, unknown>);
    }
    expect(got.length).toBe(3);
    expect((got[1] as unknown[])[0]).toEqual({ t: 's', v: 'South' });
  });

  it('collectPivotCache returns a full PivotCacheTable', async () => {
    const h = await openXlsb(pcXlsb(), { parsePivotCaches: true });
    const pc = await h.collectPivotCache('PivotCache1');
    expect(pc.rowCount).toBe(3);
    expect(pc.recordCount).toBe(3);
    expect(pc.fieldNames).toEqual(['Region']);
  });
});
