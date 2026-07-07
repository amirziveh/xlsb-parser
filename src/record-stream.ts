// Low-level XLSB record stream primitives.
// Spec ref: MS-XLSB §2.1.4 (record stream) and §2.5.122 (Brt* record types).

export interface XlsbRecord {
  type: number;
  data: Uint8Array;
}

// Shared text decoders (constructed once, reused across calls).
export const dec16 = new TextDecoder('utf-16le');
export const dec8 = new TextDecoder('utf-8');

// ---- record-type opcodes (subset we recognise) ----
export const BRT_ROW_HEADER = 0x00;
export const BRT_CELL_BLANK = 0x01;
export const BRT_CELL_RK = 0x02;
export const BRT_CELL_ERROR = 0x03;
export const BRT_CELL_BOOL = 0x04;
export const BRT_CELL_REAL = 0x05;
export const BRT_CELL_ST = 0x06;
export const BRT_CELL_ISST = 0x07;
export const BRT_FMLA_STRING = 0x08;
export const BRT_FMLA_NUM = 0x09;
export const BRT_FMLA_BOOL = 0x0a;
export const BRT_FMLA_ERROR = 0x0b;
export const BRT_SHORT_BLANK = 0x0c;
export const BRT_SHORT_RK = 0x0d;
export const BRT_SHORT_ERROR = 0x0e;
export const BRT_SHORT_BOOL = 0x0f;
export const BRT_SHORT_REAL = 0x10;
export const BRT_SHORT_ST = 0x11;
export const BRT_SHORT_ISST = 0x12;
export const BRT_SST_ITEM = 0x13;
export const BRT_BUNDLE_SH = 0x9c;
export const BRT_BUNDLE_SH_NEW = 0x0e01;

// Pivot cache record types (MS-XLSB §2.4 / §2.1.7.38, §2.1.7.39). Verified
// against real Excel .xlsb outputs.
export const BRT_BEGIN_PCD_FIELD = 0x1b81;
export const BRT_BEGIN_PCD_ATBL = 0x1e81;
export const BRT_BEGIN_PCDIRUN = 0x1f81;
export const BRT_PCDI_STRING = 0x0018;
export const BRT_PCDI_STRING2 = 0x001f;
export const BRT_PCDIDATETIME = 0x0020;
export const BRT_PCDINUMBER = 0x0015;
export const BRT_PCDIBOOLEAN = 0x0016;
export const BRT_PCDIERROR = 0x0017;
export const BRT_PCDIMISSING = 0x0014;
export const BRT_PCDIINDEX = 0x001a;
export const BRT_BEGIN_PIVOT_CACHE_RECORDS = 0x2081;
export const BRT_PC_RECORD = 0x0021;
export const BRT_PC_RECORD_DT = 0x0022;
export const BRT_END_PIVOT_CACHE_RECORDS = 0x2101;

// Excel error codes — MS-OFFBFISO §2.5.97 (BrtErr)
export const ERRORS: Record<number, string> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
  0x2b: '#GETTING_DATA',
};

// ---- record stream iterator ----
// Record header: varint type (1 or 2 bytes when bit 7 set) + varint size,
// then `size` bytes of payload.
// Throws on truncation rather than silently clamping; the parser must surface
// malformed inputs as errors, not garbage records.
export function* records(data: Uint8Array): Generator<XlsbRecord> {
  let off = 0;
  while (off < data.length) {
    const recStart = off;
    if (off >= data.length) break;
    let t = data[off++];
    if ((t & 0x80) !== 0) {
      if (off >= data.length) {
        throw new Error(
          `Truncated .bin: record type byte at offset ${recStart} announces a second byte but only ${data.length} bytes total remain`,
        );
      }
      t = ((t & 0x7f) << 7) | data[off++];
    }
    let s = 0,
      sh = 0,
      b: number;
    do {
      if (off >= data.length) {
        throw new Error(
          `Truncated .bin: record at offset ${recStart} (type 0x${t.toString(16)}) declared size varint overruns the buffer`,
        );
      }
      b = data[off++];
      s |= (b & 0x7f) << sh;
      sh += 7;
    } while (b & 0x80);
    if (off + s > data.length) {
      throw new Error(
        `Truncated .bin: record at offset ${recStart} (type 0x${t.toString(16)}) declared size ${s} but only ${data.length - off} bytes remain`,
      );
    }
    yield { type: t, data: data.subarray(off, off + s) };
    off += s;
  }
}

// ---- numeric primitives ----

export function readU16(d: Uint8Array, off: number): number {
  return d[off] | (d[off + 1] << 8);
}

export function readU32(d: Uint8Array, off: number): number {
  return (d[off] | (d[off + 1] << 8) | (d[off + 2] << 16) | (d[off + 3] << 24)) >>> 0;
}

export function readF64(d: Uint8Array, off: number): number {
  return new DataView(d.buffer, d.byteOffset + off, 8).getFloat64(0, true);
}

export function readWideString(d: Uint8Array, off: number): string {
  const len = readU32(d, off);
  return dec16.decode(d.subarray(off + 4, off + 4 + len * 2));
}

// BrtRichStr (MS-XLSB §2.5.31): 1 byte flags, then uint32 cch + chars,
// optionally followed by formatting runs and phonetic metadata. We ignore
// the runs/phonetic; the flag bit fRt only signals their *presence*, the
// cch+chars always come first.
export function readRichString(d: Uint8Array, off: number): string {
  return readWideString(d, off + 1);
}

// ---- RK number decode (MS-XLSB §2.5.122 BrtColor / Office RK encoding) ----
// RK packs a number into 4 bytes: 1 bit fx100, 1 bit fInt, 30 bits payload.
// If fInt, payload is a 30-bit signed int (interpreted as is / 100 if fx100).
// If !fInt, the payload is the high 30 bits of an IEEE 754 double.
//
// The previous implementation used BigInt + a fresh ArrayBuffer per call,
// which is the slowest possible path on V8. We pre-allocate a single scratch
// DataView and reuse it; RK cells are extremely common (Excel prefers them
// over INTEGER/REAL for compactness) so this matters for 100k-row sheets.
const RK_SCRATCH_BUF = new ArrayBuffer(8);
const RK_SCRATCH_DV = new DataView(RK_SCRATCH_BUF);
const RK_SCRATCH_U32 = new Uint32Array(RK_SCRATCH_BUF);

export function decodeRk(rk: number): number {
  const fx100 = rk & 0x01;
  const fInt = (rk >> 1) & 0x01;
  const num = rk >>> 2;
  let val: number;
  if (fInt) {
    // 30-bit signed payload, sign-extended via shift
    val = (num << 2) >> 2;
  } else {
    // High 30 bits of IEEE 754 double: bits 34..63 of the 8-byte little-endian
    // u64. The 64-bit view decodes as U32[0]=bits0..31, U32[1]=bits32..63, so
    // num occupies bits 2..31 of U32[1]. Equivalent to BigInt(num) << 34n
    // but without the per-call allocation + BigInt arithmetic.
    RK_SCRATCH_U32[0] = 0;
    RK_SCRATCH_U32[1] = num << 2;
    val = RK_SCRATCH_DV.getFloat64(0, true);
  }
  if (fx100) val /= 100;
  return val;
}

// ---- debug hex helper (used by dumpBinary) ----
export function hex(d: Uint8Array, max = 48): string {
  return Array.from(d.subarray(0, Math.min(max, d.length)))
    .map((b) => b.toString(16).padStart(2, '0'))
    .join(' ');
}
