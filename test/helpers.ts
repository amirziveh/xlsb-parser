import { zipSync, type ZipInput } from 'fflate';

// ---- primitive encoders ----

export function u32(v: number): Uint8Array {
  return new Uint8Array([v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF, (v >> 24) & 0xFF]);
}

export function u16le(s: string): Uint8Array {
  const buf = new ArrayBuffer(s.length * 2);
  const view = new Uint16Array(buf);
  for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i);
  return new Uint8Array(buf);
}

export function f64(v: number): Uint8Array {
  const buf = new ArrayBuffer(8);
  new DataView(buf).setFloat64(0, v, true);
  return new Uint8Array(buf);
}

// Record type encoding: ((b1 & 0x7F) << 7) | b2 for types >= 128
export function encType(v: number): Uint8Array {
  if (v < 128) return new Uint8Array([v]);
  return new Uint8Array([((v >> 7) & 0x7F) | 0x80, v & 0x7F]);
}

// Record size: standard 7-bit varint
export function encSize(v: number): Uint8Array {
  const bytes: number[] = [];
  do { bytes.push((v & 0x7F) | (v > 0x7F ? 0x80 : 0)); v >>>= 7; } while (v > 0);
  return new Uint8Array(bytes);
}

export function rec(type: number, data: Uint8Array): Uint8Array {
  const t = encType(type);
  const s = encSize(data.length);
  const out = new Uint8Array(t.length + s.length + data.length);
  out.set(t, 0);
  out.set(s, t.length);
  out.set(data, t.length + s.length);
  return out;
}

export function concat(...parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) { out.set(p, off); off += p.length; }
  return out;
}

// ---- cell record builders (long form: 4 col + 2 iStyleRef + 2 reserved + value) ----

export function cellIsst(col: number, sstIdx: number): Uint8Array {
  return rec(0x07, concat(u32(col), new Uint8Array(4), u32(sstIdx)));
}

export function cellReal(col: number, val: number): Uint8Array {
  return rec(0x05, concat(u32(col), new Uint8Array(4), f64(val)));
}

export function cellBlank(col: number): Uint8Array {
  return rec(0x01, concat(u32(col), new Uint8Array(4)));
}

// Variant with explicit iStyleRef bytes (low, high) at offsets 4..5
export function cellRealStyled(col: number, val: number, styleLo: number, styleHi: number): Uint8Array {
  return rec(0x05, concat(u32(col), new Uint8Array([styleLo, styleHi, 0, 0]), f64(val)));
}

// Row header: BRT_ROW_HEADER (0x00), data = uint32 rowIndex
export function rowHeader(row: number): Uint8Array {
  return rec(0x00, u32(row));
}

// ---- rich-string SST item builder ----
// format: 1 byte flags (0=no runs,0x8=has runs) + uint32 cch + chars [+ u16 cRun + _Run*4 if has runs]
export function sstItemPlain(s: string): Uint8Array {
  return rec(0x13, concat(new Uint8Array([0]), u32(s.length), u16le(s)));
}

// ---- workbook bin record (BRT_BUNDLE_SH_NEW = 0x0E01) ----
export function workbookBinRecord(sheetNames: string[]): Uint8Array {
  const parts: Uint8Array[] = [];
  sheetNames.forEach((name, i) => {
    const rId = `rId${i + 1}`;
    parts.push(rec(0x0E01, concat(
      u32(i + 1),                    // iTabID
      new Uint8Array([0, 0, 0, 0]),  // fHidden + reserved
      u32(rId.length),               // iStMeta
      u16le(rId),                    // rId
      u32(name.length),              // nameLen
      u16le(name),                   // name
    )));
  });
  return concat(...parts);
}

// ---- XLSB ZIP assembler ----

export interface XlsbParts {
  sheetNames: string[];
  sheetRecords: Uint8Array[];      // one per sheet
  sharedStrings: string[];
  extraEntries?: Record<string, Uint8Array>;
}

export function buildXlsb(parts: XlsbParts): Uint8Array {
  const text = (s: string) => new TextEncoder().encode(s);

  const overrides = parts.sheetNames.map((_, i) =>
    `<Override PartName="/xl/worksheets/sheet${i + 1}.bin" ContentType="application/vnd.ms-excel.worksheet.binary"/>`,
  ).join('');

  const contentTypes = text(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="bin" ContentType="application/vnd.ms-excel.binaryIndex"/>'
    + '<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>'
    + overrides
    + '<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles.binary"/>'
    + '</Types>',
  );

  const relsRoot = text(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>'
    + '</Relationships>',
  );

  const sheetRels = parts.sheetNames.map((_, i) =>
    `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.bin"/>`,
  ).join('');
  const relsWb = text(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + sheetRels
    + '<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>'
    + '</Relationships>',
  );

  const wbRec = workbookBinRecord(parts.sheetNames);
  const stylesRec = rec(0x0000, new Uint8Array(0));
  const ssRecs = parts.sharedStrings.length > 0
    ? concat(...parts.sharedStrings.map(sstItemPlain))
    : new Uint8Array(0);

  const zipInput: ZipInput = {
    '[Content_Types].xml': [contentTypes, { level: 0 }],
    '_rels/.rels': [relsRoot, { level: 0 }],
    'xl/_rels/workbook.bin.rels': [relsWb, { level: 0 }],
    'xl/workbook.bin': [wbRec, { level: 0 }],
    'xl/styles.bin': [stylesRec, { level: 0 }],
    ...(parts.sharedStrings.length > 0 ? { 'xl/sharedStrings.bin': [ssRecs, { level: 0 }] } : {}),
  };
  parts.sheetRecords.forEach((sr, i) => {
    zipInput[`xl/worksheets/sheet${i + 1}.bin`] = [sr, { level: 0 }];
  });
  if (parts.extraEntries) {
    for (const [k, v] of Object.entries(parts.extraEntries)) {
      zipInput[k] = [v, { level: 0 }];
    }
  }

  return new Uint8Array(zipSync(zipInput, { level: 0 }));
}

// ---- legacy entry point ----

export function makeMinimalXlsb(): Uint8Array {
  return buildXlsb({
    sheetNames: ['Sheet1'],
    sharedStrings: ['Name', 'Value', 'foo', 'bar'],
    sheetRecords: [concat(
      rowHeader(0), cellIsst(0, 0), cellIsst(1, 1),
      rowHeader(1), cellIsst(0, 2), cellReal(1, 42.5),
      rowHeader(2), cellIsst(0, 3), cellBlank(1),
    )],
  });
}
