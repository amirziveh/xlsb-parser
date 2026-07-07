import {
  records,
  readU32,
  readWideString,
  BRT_BUNDLE_SH,
  BRT_BUNDLE_SH_NEW,
} from './record-stream.js';

// Parse xl/workbook.bin → ordered list of sheet names.
// BrtBundleSh (0x9C) is the legacy form; BrtBundleShNew (0x0E01) used by XLSB.
// Both share the layout: iTabID(4) + fHidden(1) + reserved(3) + iStMeta(4)
// + rId chars + nameLen(4) + name chars.
export function parseWorkbook(data: Uint8Array): string[] {
  const names: string[] = [];
  for (const r of records(data)) {
    if (r.type === BRT_BUNDLE_SH_NEW && r.data.length >= 12) {
      const rIdLen = readU32(r.data, 8);
      const nameOff = 12 + rIdLen * 2;
      if (nameOff + 4 <= r.data.length) names.push(readWideString(r.data, nameOff));
    } else if (r.type === BRT_BUNDLE_SH && r.data.length >= 9) {
      const rIdLen = readU32(r.data, 5);
      const nameOff = 9 + rIdLen * 2;
      if (nameOff + 4 <= r.data.length) names.push(readWideString(r.data, nameOff));
    }
  }
  return names;
}
