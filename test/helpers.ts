import { unzipSync, zipSync, type ZipInput } from 'fflate';

const dec16 = new TextDecoder('utf-16le');

function u32(v: number): Uint8Array {
  return new Uint8Array([v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF, (v >> 24) & 0xFF]);
}

function f64(v: number): Uint8Array {
  const buf = new ArrayBuffer(8);
  new DataView(buf).setFloat64(0, v, true);
  return new Uint8Array(buf);
}

function u16le(s: string): Uint8Array {
  const buf = new ArrayBuffer(s.length * 2);
  const view = new Uint16Array(buf);
  for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i);
  return new Uint8Array(buf);
}

// Record type encoding: ((b1 & 0x7F) << 7) | b2 for types >= 128
function encType(v: number): Uint8Array {
  if (v < 128) return new Uint8Array([v]);
  return new Uint8Array([((v >> 7) & 0x7F) | 0x80, v & 0x7F]);
}

// Record size: standard 7-bit varint
function encSize(v: number): Uint8Array {
  const bytes: number[] = [];
  do { bytes.push((v & 0x7F) | (v > 0x7F ? 0x80 : 0)); v >>>= 7; } while (v > 0);
  return new Uint8Array(bytes);
}

function rec(type: number, data: Uint8Array): Uint8Array {
  const t = encType(type);
  const s = encSize(data.length);
  const out = new Uint8Array(t.length + s.length + data.length);
  out.set(t, 0);
  out.set(s, t.length);
  out.set(data, t.length + s.length);
  return out;
}

export function makeMinimalXlsb(): Uint8Array {
  const text = (s: string) => new TextEncoder().encode(s);

  const xmlHeaders = {
    contentTypes: text(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="bin" ContentType="application/vnd.ms-excel.binaryIndex"/>'
      + '<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>'
      + '<Override PartName="/xl/worksheets/sheet1.bin" ContentType="application/vnd.ms-excel.worksheet.binary"/>'
      + '<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles.binary"/>'
      + '</Types>',
    ),
    relsRoot: text(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>'
      + '</Relationships>',
    ),
    relsWb: text(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/>'
      + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>'
      + '</Relationships>',
    ),
  };

  // workbook.bin: BRT_BUNDLE_SH_NEW (0x0E01)
  // Record format:
  //   4 bytes: iTabID
  //   1 byte:  fHidden
  //   3 bytes: reserved
  //   4 bytes: iStMeta (rId char length)
  //   N*2 bytes: rId string (UTF-16LE)
  //   4 bytes: nameLen
  //   nameLen*2 bytes: sheet name (UTF-16LE)
  const sheetName = 'Sheet1';
  const rId = 'rId1';
  const wbData = concat(
    u32(1),                    // iTabID
    new Uint8Array([0, 0, 0, 0]), // fHidden(1) + reserved(3)
    u32(rId.length),           // iStMeta at offset 8
    u16le(rId),                // rId at offset 12
    u32(sheetName.length),     // nameLen at 12 + rIdLen*2 = 20
    u16le(sheetName),          // name at 24
  );
  const wbRec = rec(0x0E01, wbData);

  // styles.bin: empty
  const stylesRec = rec(0x0000, new Uint8Array(0));

  // Cell record format (long form):
  //   4 bytes: col (uint32)
  //   2 bytes: iStyleRef (uint16, sign-extended)
  //   2 bytes: reserved/fPhShow
  //   value data follows at offset 8
  // Cell types:
  //   BRT_CELL_ISST (0x07): value = 4 bytes SST index
  //   BRT_CELL_REAL (0x05): value = 8 bytes IEEE 754 double
  //   BRT_CELL_BLANK (0x01): no value

  function cellIsst(col: number, sstIdx: number): Uint8Array {
    return rec(0x07, concat(u32(col), new Uint8Array(4), u32(sstIdx)));
  }

  function cellReal(col: number, val: number): Uint8Array {
    return rec(0x05, concat(u32(col), new Uint8Array(4), f64(val)));
  }

  function cellBlank(col: number): Uint8Array {
    return rec(0x01, concat(u32(col), new Uint8Array(4)));
  }

  // Row header: BRT_ROW_HEADER (0x00)
  //   data = uint32 rowIndex
  function rowHeader(row: number): Uint8Array {
    return rec(0x00, u32(row));
  }

  const sheetRecs = concat(
    // Row 0: headers
    rowHeader(0),
    cellIsst(0, 0), // A1 → "Name" (SST idx 0)
    cellIsst(1, 1), // B1 → "Value" (SST idx 1)

    // Row 1: data
    rowHeader(1),
    cellIsst(0, 2), // A2 → "foo" (SST idx 2)
    cellReal(1, 42.5), // B2 → 42.5

    // Row 2: blank
    rowHeader(2),
    cellIsst(0, 3), // A3 → "bar" (SST idx 3)
    cellBlank(1),   // B3 → blank
  );

  // sharedStrings.bin: BRT_SST_ITEM (0x13) records
  //   format: 1 byte flags + 4 bytes charCount + charCount*2 bytes UTF-16LE
  const sharedStrings = ['Name', 'Value', 'foo', 'bar'];
  const ssRecs = concat(...sharedStrings.map(s =>
    rec(0x13, concat(new Uint8Array([0]), u32(s.length), u16le(s))),
  ));

  const zipInput: ZipInput = {
    '[Content_Types].xml': [xmlHeaders.contentTypes, { level: 0 }],
    '_rels/.rels': [xmlHeaders.relsRoot, { level: 0 }],
    'xl/_rels/workbook.bin.rels': [xmlHeaders.relsWb, { level: 0 }],
    'xl/workbook.bin': [wbRec, { level: 0 }],
    'xl/styles.bin': [stylesRec, { level: 0 }],
    'xl/worksheets/sheet1.bin': [sheetRecs, { level: 0 }],
    'xl/sharedStrings.bin': [ssRecs, { level: 0 }],
  };

  return new Uint8Array(zipSync(zipInput, { level: 0 }));
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) { out.set(p, off); off += p.length; }
  return out;
}
