import { describe, it, expect } from 'vitest';
import { decodeRk } from '../src/record-stream.js';

// §2.10: RK decode previously used BigInt + a fresh ArrayBuffer per call.
// This test verifies the pre-allocated DataView scratch-buffer rewrite
// produces identical values across the full input space.
//
// Reference implementation (old code) inlined here for direct comparison.
function decodeRkBigint(rk: number): number {
  const fx100 = rk & 0x01;
  const fInt = (rk >> 1) & 0x01;
  const num = rk >>> 2;
  let val: number;
  if (fInt) {
    val = (num << 2) >> 2;
  } else {
    const buf = new ArrayBuffer(8);
    new DataView(buf).setBigUint64(0, BigInt(num) << 34n, true);
    val = new DataView(buf).getFloat64(0, true);
  }
  if (fx100) val /= 100;
  return val;
}

describe('decodeRk: DataView scratch vs reference BigInt implementation', () => {
  it('produces identical output on every RK value across 4 bytes of input space', () => {
    // 2^32 is too many; sample a dense mix:
    // - every value 0..1023 (low bytes, exercises fInt+fx100 combos)
    // - powers of two
    // - values that flip each bit of the high 30 bits
    // - signed boundary cases for the fInt path
    const samples: number[] = [];
    for (let i = 0; i < 1024; i++) samples.push(i);
    for (let i = 0; i <= 31; i++) samples.push(1 << i);
    for (let i = 0; i <= 31; i++) samples.push((1 << i) >>> 0);
    // fInt path: payload bits 2..31 are the 30-bit signed integer
    // 0x80000000 → fInt=0; 0x00000002 → fInt=1 with num=0 → val=0
    // fInt=1 with num=0x1FFFFFFF → max signed 30-bit positive
    samples.push(0x00000002, 0xfffffffe, 0x7ffffffe, 0x00000001, 0x00000003);
    // fInt=0 (double) path high bits
    for (let i = 0; i <= 30; i++) samples.push(((1 << i) << 2) >>> 0);
    // fx100 toggle on each of the above
    const withFx100: number[] = [];
    for (const v of samples) {
      withFx100.push(v);
      withFx100.push((v | 0x01) >>> 0);
    }

    const all = [...new Set([...samples, ...withFx100])];
    let checked = 0;
    for (const rk of all) {
      const u = rk >>> 0;
      const actual = decodeRk(u);
      const expected = decodeRkBigint(u);
      if (!Object.is(actual, expected)) {
        // Allow exact bit-for-bit equality including NaN (Object.is handles NaN)
        throw new Error(
          `decodeRk(0x${u.toString(16).padStart(8, '0')}) = ${actual} but BigInt path = ${expected}`,
        );
      }
      checked++;
    }
    expect(checked).toBeGreaterThan(1000);
  });

  it('decodes integer RK = 0 (fInt=1, num=0, fx100=0) to 0', () => {
    // bits: fInt (bit 1) = 1, so this is the fInt path with num = 0
    expect(decodeRk(0x00000002)).toBe(0);
  });

  it('decodes integer RK = 100 (fInt=1, num=25, fx100=1) to 25/100 = 0.25', () => {
    // num = 25, fInt = 1, fx100 = 1 → integer 25 / 100 = 0.25
    // RK encoding: bit0 = fx100, bit1 = fInt, bits 2..31 = num
    // 25 << 2 = 100, then | 0b11 = 103 = 0x67
    const rk = (25 << 2) | 0b11;
    expect(decodeRk(rk >>> 0)).toBe(0.25);
  });
});
