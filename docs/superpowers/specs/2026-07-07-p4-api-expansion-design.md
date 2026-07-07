# P4 — API Surface Expansion Design

**Date:** 2026-07-07
**Goal:** Promote `xlsb-parser` from a buffered prototype to a world-level library by adding an options bag, streaming iteration, styles/date metadata, and making pivot caches opt-in.

**Approved decisions (from user, 2026-07-07):**

1. **Options bag as 2nd arg** — `parseXlsb(data, options?)`. Clean break from the legacy `parseXlsb(data, onProgress?)` signature. Deprecation cycle: detect a function passed as the 2nd arg and treat it as `onProgress` for backwards compatibility through 1.x; remove at 2.0.
2. **`openXlsb()` handle pattern for streaming** — `const handle = await openXlsb(data); for await (const row of handle.iterSheetRows(0)) { ... }`. Unzips once, returns handle exposing `sheetNames: string[]` and `iterSheetRows(index): AsyncGenerator<ParsedRow>`. Memory: O(1) per row instead of O(N).
3. **Expose date metadata; no auto-convert** — `cell.numFmtId?: number`, `cell.isDate?: boolean`, `cell.dateValue?: string` (ISO 8601). `cell.v` remains the raw serial number. Consumer chooses whether to read `dateValue` or compute their own.
4. **Pivot caches opt-in** — `parsePivotCaches: false` by default in the options bag. Heuristic decoder stays for callers who explicitly opt in. Honest "Limitations" section added to README.

---

## Architecture

### Public API (after P4)

```ts
// New signature (1.x): options bag as 2nd arg.
// 2nd arg = function → treated as onProgress (deprecated, works through 1.x).
export interface ParseOptions {
  onProgress?: ProgressCallback;
  maxZipBytes?: number;          // default unset; throws if decompressed > N
  maxRowsPerSheet?: number;      // default unset; stops parsing a sheet at N rows
  dumpBinaries?: boolean;        // default false (was: always on). Opt-in debug dumps.
  readXml?: boolean;             // default false (was: always on). Opt-in raw XML.
  parsePivotCaches?: boolean;    // default false (was: always on). Opt-in.
}

export function parseXlsb(
  data: ArrayBuffer | Uint8Array,
  options?: ParseOptions,
): Promise<ParsedXlsb>;

// Streaming handle for huge sheets — unzips once, exposes iterators.
export interface XlsbHandle {
  sheetNames: string[];
  sharedStrings: string[];
  styles: StylesTable | null;
  iterSheetRows(sheetIndex: number, options?: IterOptions): AsyncGenerator<ParsedRow>;
}

export interface IterOptions {
  maxRows?: number;               // stop after N rows
  onProgress?: ProgressCallback;
}

export function openXlsb(
  data: ArrayBuffer | Uint8Array,
  options?: ParseOptions,
): Promise<XlsbHandle>;

// Augmented Cell type — new optional fields, no breaking changes.
export interface Cell {
  t: 'n' | 's' | 'b' | 'e' | 'blank' | 'f';
  v?: number | string | boolean;
  err?: string;
  ixf?: number;              // existing: signed iStyleRef
  numFmtId?: number;         // new: numeric format ID resolved from styles
  isDate?: boolean;          // new: true when numFmtId is a date/time format
  dateValue?: string;        // new: ISO 8601 string when isDate && t === 'n'
}
```

### Components

| Module | Responsibility | New in P4? |
|---|---|---|
| `src/types.ts` | Public interfaces. Add `ParseOptions`, `XlsbHandle`, `IterOptions`. Extend `Cell` with `numFmtId`, `isDate`, `dateValue`. | yes |
| `src/record-stream.ts` | `records()` iterator + numeric primitives + RK decode. | no |
| `src/workbook.ts` | `parseWorkbook`. | no |
| `src/shared-strings.ts` | `parseSharedStrings`. | no |
| `src/sheet.ts` | `parseSheet` (buffered) — accepts optional `maxRows` and a `StylesTable` for date metadata. New: `iterSheet` generator variant. | extend |
| `src/styles.ts` | **NEW.** Parses `xl/styles.bin` → `StylesTable { cellXfs: NumFmtId[], numFmts: Map<number, string> }`. Includes `isDateFormatId(id)` heuristic. | yes |
| `src/dump.ts` | `dumpBinary` (debug). Only invoked when `dumpBinaries: true`. | conditional |
| `src/pivot-cache.ts` | `parsePivotCache` (heuristic). Only invoked when `parsePivotCaches: true`. No rewrite. | conditional |
| `src/index.ts` | Orchestrator. Rewrite `parseXlsb` to take `ParseOptions`. Add `openXlsb`. | yes |

### Data flow

`parseXlsb(data, opts)`:
1. Unzip. If `maxZipBytes` set and decompressed total exceeds it, throw `XlsbSizeError`.
2. Parse workbook, shared strings, **styles** (new).
3. For each sheet: parse via `parseSheet(bytes, ss, { maxRows, styles })`.
4. Each cell gets `numFmtId` from `styles.cellXfs[ixf]`, `isDate` from styles table, `dateValue` if applicable.
5. `binaryDumps` only populated if `dumpBinaries: true`.
6. `xmlFiles` only populated if `readXml: true`.
7. `pivotCaches` only populated if `parsePivotCaches: true`.

`openXlsb(data, opts)`:
1. Same unzip + parse workbook + shared strings + styles.
2. Returns handle. Does NOT eagerly parse sheets.
3. `handle.iterSheetRows(i)` yields rows one at a time from the `records()` generator without buffering them into an array.

### Date detection

Per OOXML: cell style `iStyleRef` (already exposed as `ixf`) is an index into the `cellXfs` table in `styles.bin`. Each `cellXfs[i]` has a `numFmtId`. The `numFmt` string determines whether a cell is a date:
- Built-in date format IDs (14–22, 27–36, 45–47, 50–58, 78–81) — Excel's standard date/time formats.
- Custom `numFmt` strings containing `d`, `m`, `y`, `h`, `s` (date/time tokens) outside escaped/literal sections.

For serial-to-ISO conversion we use Excel's 1900 epoch: `Date(1899, 11, 30) + serial * 86400_000ms`; produce ISO 8601 `YYYY-MM-DDTHH:mm:ss.sssZ`.

### Error handling

- `XlsbSizeError` (subclass of `Error`) thrown when `maxZipBytes` or `maxRowsPerSheet` exceeded. Distinct from parse errors so consumers can differentiate.
- Truncation errors from `records()` (introduced in P1) unchanged.

### Testing strategy

- TDD each slice: write failing test → implement → green.
- Backwards-compat: keep all 56 existing tests green; add new tests for the options-bag signature.
- Streaming: test `openXlsb` with a 1000-row synthetic sheet — assert memory stays low and all rows yielded.
- Styles/dates: synthesize a `styles.bin` fixture with known numFmtIds, assert `cell.numFmtId`/`isDate`/`dateValue` propagate.
- Pivot opt-in: assert `pivotCaches: []` when option absent, populated when `parsePivotCaches: true`.

### Out of scope (deferred)

- Spec-driven pivot-cache rewrite (was P4 in the original audit; deferred to post-1.0).
- Streaming the unzip step itself (would require replacing `fflate` `unzipSync` with a zip-stream library; deferred).
- Web-worker offload of `parseXlsb` (deferred; the streaming API covers the main pain).
