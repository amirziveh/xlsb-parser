import { describe, it, expect } from 'vitest';
import { parseXlsb } from '../src/index.js';
import {
  buildXlsb, pcdFieldFull, pcdStr, pcdDate, f64,
  pcRecordsHeader, pcRecord, pcRecordsEnd, concat, u32,
} from './helpers';

function section38Def(): Uint8Array {
  return concat(
    pcdFieldFull('CustomerName', { isSrc: true, fText: true }),
    pcdStr('Great Lakes Food Market'),
    pcdStr('Richter Supermarkt'),
    pcdFieldFull('OrderDate', { isSrc: true, fDate: true }),
    pcdDate(1997, 5, 6, 0, 0, 0),
    pcdFieldFull('ProductName', { isSrc: true, fText: true }),
    pcdStr('Geitost'),
    pcdStr('Gnocchi di nonna Alice'),
    pcdFieldFull('UnitPrice', { isSrc: true, fNum: true }),
    pcdFieldFull('Quantity', { isSrc: true, fNum: true }),
  );
}

function section38Recs(): Uint8Array {
  return concat(
    pcRecordsHeader(1),
    pcRecord([u32(0), u32(0), u32(0), f64(2.5), f64(8)]),
    pcRecordsEnd(),
  );
}

describe('MS-XLSB §3.8 worked example', () => {
  it('decodes a 5-field cache with indexed string, date, and number fields', async () => {
    const xlsb = buildXlsb({
      sheetNames: ['S'], sharedStrings: [], sheetRecords: [],
      extraEntries: {
        'xl/pivotCache/pivotCacheDefinition1.bin': section38Def(),
        'xl/pivotCache/pivotCacheRecords1.bin': section38Recs(),
      },
    });
    const wb = await parseXlsb(xlsb, { parsePivotCaches: true });
    const pc = wb.pivotCaches[0];
    expect(pc.fieldNames).toEqual(['CustomerName', 'OrderDate', 'ProductName', 'UnitPrice', 'Quantity']);
    expect(pc.fields.map(f => f.kind)).toEqual(['indexed', 'date', 'indexed', 'number', 'number']);
    expect(pc.fields[0].sharedItems.length).toBe(2);
    expect(pc.fields[0].sharedItems[0]).toEqual({ t: 's', v: 'Great Lakes Food Market' });
    expect(pc.fields[1].sharedItems[0]).toEqual({ t: 'd', v: '1997-05-06' });
    expect(pc.rows.length).toBe(1);
    expect(pc.rows[0]).toEqual([
      { t: 's', v: 'Great Lakes Food Market' },
      { t: 'd', v: '1997-05-06' },
      { t: 's', v: 'Geitost' },
      { t: 'n', v: 2.5 },
      { t: 'n', v: 8 },
    ]);
  });
});
