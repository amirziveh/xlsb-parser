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
  workbookBin?: Uint8Array;        // override the default workbook record
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

  const wbRec = parts.workbookBin ?? workbookBinRecord(parts.sheetNames);
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

// ---- styles.bin builder ----
// Builds a minimal xl/styles.bin with one cellXfs entry whose numFmtId is set.
// BrtFmt (0x02D9) records (custom formats) precede BrtCellXF (0x01F9) records.
// BrtFmt data: uint16 numFmtId + uint32 fmtStrLen + UTF-16LE chars.
// BrtCellXF data: uint16 numFmtId at byte 0 (+ more bytes we don't care about).
export function stylesBinRecord(opts: {
  cellXfs?: number[];                // numFmtId for each cellXfs entry
  customFmts?: Record<number, string>; // numFmtId → format string
}): Uint8Array {
  const parts: Uint8Array[] = [];
  // Custom formats
  if (opts.customFmts) {
    for (const [idStr, fmt] of Object.entries(opts.customFmts)) {
      const id = Number(idStr);
      parts.push(rec(0x02D9, concat(
        new Uint8Array([id & 0xFF, (id >> 8) & 0xFF]),  // uint16 numFmtId
        u32(fmt.length),
        u16le(fmt),
      )));
    }
  }
  // cellXfs — one BrtCellXF per entry, numFmtId at byte 0
  if (opts.cellXfs) {
    for (const numFmtId of opts.cellXfs) {
      // Pad to 4 bytes minimum (BrtCellXF has more fields but we only set
      // the numFmtId; the rest are zero):
      parts.push(rec(0x01F9, new Uint8Array([numFmtId & 0xFF, (numFmtId >> 8) & 0xFF, 0, 0])));
    }
  }
  return concat(...parts);
}

// Build a cell with explicit iStyleRef bytes (so it points to a specific
// cellXfs entry in our synthesized styles table).
export function cellRealWithStyle(col: number, val: number, styleLo: number, styleHi: number): Uint8Array {
  return rec(0x05, concat(u32(col), new Uint8Array([styleLo, styleHi, 0, 0]), f64(val)));
}

// ---- comprehensive cell builders (long form: 4 col + 2 iStyleRef + 2 reserved + value) ----
// Cell record opcodes per MS-XLSB §2.4
const BRT_CELL_RK_OP = 0x02;
const BRT_CELL_ERROR_OP = 0x03;
const BRT_CELL_BOOL_OP = 0x04;
const BRT_CELL_ST_OP = 0x06;
const BRT_FMLA_STRING_OP = 0x08;
const BRT_FMLA_NUM_OP = 0x09;
const BRT_FMLA_BOOL_OP = 0x0A;
const BRT_FMLA_ERROR_OP = 0x0B;

// Build a long-form cell with explicit type, col and value bytes
export function cellLong(op: number, col: number, valueBytes: Uint8Array): Uint8Array {
  return rec(op, concat(u32(col), new Uint8Array(4), valueBytes));
}

export function cellRk(col: number, rkU32: number): Uint8Array {
  return cellLong(BRT_CELL_RK_OP, col, u32(rkU32));
}

export function cellError(col: number, errCode: number): Uint8Array {
  return cellLong(BRT_CELL_ERROR_OP, col, new Uint8Array([errCode]));
}

export function cellBool(col: number, v: boolean): Uint8Array {
  return cellLong(BRT_CELL_BOOL_OP, col, new Uint8Array([v ? 1 : 0]));
}

export function cellStringInline(col: number, s: string): Uint8Array {
  return cellLong(BRT_CELL_ST_OP, col, concat(new Uint8Array([0]), u32(s.length), u16le(s)));
}

export function cellFmlaNum(col: number, v: number): Uint8Array {
  return cellLong(BRT_FMLA_NUM_OP, col, f64(v));
}

export function cellFmlaString(col: number, s: string): Uint8Array {
  return cellLong(BRT_FMLA_STRING_OP, col, concat(u32(s.length), u16le(s)));
}

export function cellFmlaBool(col: number, v: boolean): Uint8Array {
  return cellLong(BRT_FMLA_BOOL_OP, col, new Uint8Array([v ? 1 : 0]));
}

export function cellFmlaError(col: number, errCode: number): Uint8Array {
  return cellLong(BRT_FMLA_ERROR_OP, col, new Uint8Array([errCode]));
}

// ---- short-form cell builders ----
// Short records: implicit col=prevCol+1, 2 bytes (unknown) + 2 bytes iStyleRef + value.
// Code reads: ixf = readU16(d, 2); readShortCell(r.type, d, 4, ss); so the value
// starts at offset 4. The leading 4 bytes are 2 unknown + 2 iStyleRef.
const BRT_SHORT_BLANK_OP = 0x0C;
const BRT_SHORT_RK_OP = 0x0D;
const BRT_SHORT_ERROR_OP = 0x0E;
const BRT_SHORT_BOOL_OP = 0x0F;
const BRT_SHORT_REAL_OP = 0x10;
const BRT_SHORT_ST_OP = 0x11;
const BRT_SHORT_ISST_OP = 0x12;

function shortRec(op: number, valueBytes: Uint8Array = new Uint8Array(0)): Uint8Array {
  // 4 bytes header (col? iStyleRef? not fully documented but tests confirm
  // code reads ixf from offset 2 and value from offset 4)
  return rec(op, concat(new Uint8Array(4), valueBytes));
}

export function shortBlank(): Uint8Array { return shortRec(BRT_SHORT_BLANK_OP); }
export function shortRk(rkU32: number): Uint8Array { return shortRec(BRT_SHORT_RK_OP, u32(rkU32)); }
export function shortError(errCode: number): Uint8Array { return shortRec(BRT_SHORT_ERROR_OP, new Uint8Array([errCode])); }
export function shortBool(v: boolean): Uint8Array { return shortRec(BRT_SHORT_BOOL_OP, new Uint8Array([v ? 1 : 0])); }
export function shortReal(v: number): Uint8Array { return shortRec(BRT_SHORT_REAL_OP, f64(v)); }
export function shortString(s: string): Uint8Array {
  return shortRec(BRT_SHORT_ST_OP, concat(new Uint8Array([0]), u32(s.length), u16le(s)));
}
export function shortIsst(sstIdx: number): Uint8Array { return shortRec(BRT_SHORT_ISST_OP, u32(sstIdx)); }

// ---- rich-text SST item ----
// BrtRichStr (MS-XLSB §2.5.31):
//   1 byte flags (bit 3 = fRt, has formatting runs; bit 7 = fPh-displayed... we set fRt)
//   uint32 cch (char count)
//   cch * 2 bytes UTF-16LE chars
//   if fRt: uint16 cRun + 4 bytes per run (we emit 0 runs to keep test focused on string)
// We expose a variant WITH fRt=1, cRun=0 to verify runs are skipped correctly
// (regression for the §2.5 audit claim that was unverified).
export function sstItemRich(s: string): Uint8Array {
  // flags = 0x08 means fRt=1 (rich formatting), but we emit 0 runs.
  // This exercises the code path where runs ARE present but text reads
  // identically to the no-runs form.
  return rec(0x13, concat(
    new Uint8Array([0x08]),       // fRt=1
    u32(s.length),
    u16le(s),
    u32(0),                        // cRun = 0
  ));
}

// Build SST bytes from a mixed list of plain/rich string specs
export function sstBytes(items: { s: string; rich?: boolean }[]): Uint8Array {
  return concat(...items.map(it => it.rich ? sstItemRich(it.s) : sstItemPlain(it.s)));
}

// ---- legacy workbook form (BRT_BUNDLE_SH = 0x9C) ----
// Layout per parseWorkbook: iTabID(4) + fHidden(1) + iStMeta(4) = 9-byte header,
// then rId chars, then nameLen(4) + name chars.
export function workbookBinRecordLegacy(sheetNames: string[]): Uint8Array {
  const parts: Uint8Array[] = [];
  sheetNames.forEach((name, i) => {
    const rId = `rId${i + 1}`;
    parts.push(rec(0x9C, concat(
      u32(i + 1),                       // iTabID (4 bytes, offset 0..3)
      new Uint8Array([0]),              // fHidden (1 byte, offset 4)
      u32(rId.length),                 // iStMeta (4 bytes, offset 5..8)
      u16le(rId),                       // rId chars (offset 9..)
      u32(name.length),
      u16le(name),
    )));
  });
  return concat(...parts);
}

