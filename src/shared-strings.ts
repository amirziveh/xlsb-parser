import { records, readRichString, BRT_SST_ITEM } from './record-stream.js';

// Parse xl/sharedStrings.bin → flat list of SST items (string only, runs
// and phonetic metadata discarded per BrtRichStr §2.5.31).
export function parseSharedStrings(data: Uint8Array): string[] {
  const list: string[] = [];
  for (const r of records(data)) {
    if (r.type === BRT_SST_ITEM && r.data.length >= 5) list.push(readRichString(r.data, 0));
  }
  return list;
}
