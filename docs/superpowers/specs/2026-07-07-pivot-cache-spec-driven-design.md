# Pivot-cache decoder: spec-driven rewrite + streaming (P3)

**Date:** 2026-07-07
**Status:** Approved design (pending implementation plan)
**Scope:** Non-OLAP PivotCaches only (Excel internal / cell-range data sources). OLAP,
SQL/external-connection, slicer/timeline, and server-formatting (`PCDISrvFmt`) caches
stay best-effort as today.

## Why

The original `src/pivot-cache.ts` was an intentional heuristic (see its own module
NOTE) — it guessed each field's type by byte-sampling the record rows instead of
reading the authoritative `BrtBeginPCDFAtbl` flags from the definition part. An audit
against `[MS-XLSB]-251113 2.docx` found the following real gaps (F1, the "wrong opcode"
claim, was a false alarm — the parser's varint decoder reads the spec's LEB128 bytes
correctly; both real dashboard files decode fully into `BrtPCRRecord` rows):

- **F2** `BrtPCRRecordDt` (per-field-value records) silently skipped.
- **F3** Field type guessed from bytes, not from `fNumField`/`fDateInField`/`fHasTextItem`.
- **F4** `PCDIDateTime.day` read as u16 → corrupts any datetime with `hr != 0`.
- **F5** `BrtBeginPCDIRun` only handles `mdSxoper` 0x02/0x20; missing 0x01 (number runs),
  0x10 (error runs); no single `BrtPCDINumber`/`BrtPCDIIndex`/`BrtPCDIMissing`/`BrtPCDIBoolean`.
- **F7** Inline date fields (`fDateInField=1`, no shared items) reinterpreted as IEEE-754 doubles.
- **F8** String detection requires `[A-Za-z]` → rejects numeric-as-text codes and all non-Latin text.
- **F9** `f64`-heuristic false positives cascade and desync every subsequent field in the row.
- **F10** `fieldNames` truncated from first 5 rows → drops sparse trailing fields.
- **F11** Robustness: `parsePivotCache` unguarded (one bad cache aborts the whole parse);
  `crecords` not validated; ad-hoc `< 100`/`< 500` length bounds silently drop legit fields.

Goal: a **spec-driven** decoder that fixes F2–F11, keeps the current `fieldNames` /
`rowCount` shape on existing files (rows may improve), adds a streaming API for the
136k-row caches seen in the real dashboards, and is validated by synthetic + §3.8-spec
fixtures (no real user file committed).

## Constraints / agreed decisions

- Scope: **Non-OLAP only**.
- API: **same opt-in flag** (`parsePivotCaches: true`), richer output.
- Back-compat: **field names match, rows can improve** (wrong values today become correct).
- Fixtures: **no trace of the real file** — pure synthetic per-branch + MS-XLSB §3.8 example.
- Scale: **streaming generator on `openXlsb`** for rows.
- Cell output: **per-cell discriminated union** (`PivotCacheCell`).
- `openXlsb` with `parsePivotCaches: true`: **eager definitions, lazy rows**.

## §1 Data model & types (`src/types.ts`)

```ts
// Per-cell value, mirrors Sheet's `Cell` shape.
export type PivotCacheCell =
  | { t: 's'; v: string }
  | { t: 'n'; v: number }
  | { t: 'd'; v: string; serial?: number }   // date-time, ISO-8601; serial optional
  | { t: 'b'; v: boolean }
  | { t: 'e'; v: string }                      // '#REF!' etc. (from BrtPCDIError)
  | { t: 'blank' };                            // BrtPCDIMissing

// Per-field descriptor built from the definition part.
export interface PivotCacheField {
  name: string;
  // How each row item for this field is encoded in BrtPCRRecord.rgb / BrtPCRRecordDt:
  kind: 'indexed' | 'number' | 'date' | 'string' | 'boolean' | 'error' | 'blank';
  sharedItems: (PivotCacheCell | null)[];      // populated when kind === 'indexed'
}

export interface PivotCacheTable {
  name: string;
  fieldNames: string[];                        // UNCHANGED shape (back-compat)
  fields: PivotCacheField[];                   // NEW: authoritative per-field info
  rows: PivotCacheCell[][];                    // CHANGED element type (still indexable)
  rowCount: number;                            // from BrtBeginPivotCacheRecords crecords
  recordCount: number;                         // actual records decoded
}
```

`PivotCacheCell[][]` is still indexable as `rows[r][c]`, so existing `.rows` consumers
keep working; they just receive a richer cell. This satisfies "field names match, rows
can improve".

## §2 Decoder core (`src/pivot-cache.ts`)

Two passes, both driven by the spec (not sampling).

### Pass 1 — Definition
Walk `records(def)`:
- `0x1B81` `BrtBeginPCDField` → open a field, read `stFldName` at offset 20 (verified layout
  on real files: 20-byte fixed header, then `XLWideString` u32 length + UTF-16LE).
- `0xBD` run-header (`BrtBeginPCDFAtbl`) → read flags byte; set `fNumField`,
  `fDateInField`, `fHasTextItem`, `citems` (clamp `citems` to spec max 1,048,576 per the
  `bVerCacheCreated >= 3` rule). Decide `field.kind`:
  - `fNumField` → `'number'`;
  - else `fDateInField && !fHasTextItem` → `'date'`;
  - else if `fHasTextItem || fTextEtcField` → `'string'` (→ `'indexed'` once shared items are seen);
  - else `'string'`.
- Shared-item records + `BrtBeginPCDIRun` (`mdSxoper` 0x01/0x02/0x10/0x20) → append to
  `field.sharedItems` (FIX F5). On `BrtEndPCDField` close the field.
- Genuine shared items present ⇒ `field.kind = 'indexed'`.

### Pass 2 — Records
Walk `records(recs)`:
- `0x2081` `BrtBeginPivotCacheRecords` → read `crecords` (4-byte u32) → `rowCount` (FIX F11).
- `0x21` `BrtPCRRecord` → for each field, consume rgb item by `field.kind`:
  - `indexed` → 4-byte u32 index → `sharedItems[idx]` (or `null` if out of range);
  - `number` → 8-byte `Xnum` (IEEE-754 double);
  - `date` → 8-byte `PCDIDateTime` → ISO-8601 string (FIX F4: `yr` u16 @+0, `mon` u16 @+2,
    `dom` **1 byte** @+4, `hr` @+5, `min` @+6, `sec` @+7; time components included in the
    ISO string only when non-zero);
  - `string` → `XLWideString` (4-byte len + UTF-16LE), **no alpha requirement** (FIX F8);
  - `boolean` / `error` / `blank` as appropriate.
- `0x22` `BrtPCRRecordDt` → per-field `BrtPCDI*` records dispatched to the matching field
  (FIX F2).
- `0x2101` `BrtEndPivotCacheRecords` → stop.

### Guards (FIX F11)
- `index.ts` wraps the `parsePivotCache` call in try/catch so a single malformed cache never
  aborts workbook parsing (mirrors the existing `parseStyles` guard).
- `citems` clamped to spec max instead of `< 100`; no `< 500` string-length cap that drops
  legitimate large fields.
- `rowCount` (`crecords`) validated against actual decoded records; mismatch logged, not thrown.
- **FIX F10:** `fieldNames` length is taken directly from the definition field count; the
  first-5-rows truncation heuristic is removed.

## §3 Streaming API (`src/handle.ts`)

`openXlsb(data, { parsePivotCaches: true })` eagerly parses cache **definitions** (Pass 1)
into `handle.pivotCaches: { name, fields, fieldNames, rowCount }[]`. Each entry holds a
lazy reader over the matching records part (kept in the already-in-memory zip buffer).

New method:
```ts
async *iterPivotCacheRows(
  indexOrName: number | string,
  opts?: { maxRows?: number; onProgress?: ProgressCallback },
): AsyncGenerator<PivotCacheCell[]> {
  // Streams BrtPCRRecord / BrtPCRRecordDt rows, reusing Pass-2 dispatch.
  // Yields one PivotCacheCell[] per source row. O(cells-per-row) decode memory.
}
```
Plus `collectPivotCache(indexOrName): Promise<PivotCacheTable>` (eager collect, mirrors
`collectSheet`). The records-part `Uint8Array` is already fully in memory (the zip is
decoded up front in `openXlsb`), so streaming means O(cells-per-row) decode memory, not
O(total rows) — consistent with `iterSheetRows`.

## §4 Fixtures & tests (no real file committed)

Per privacy constraint, fixtures are **pure synthetic** + the **MS-XLSB §3.8 worked example**:

1. `test/fixtures/pivot/` — byte-exact `pivotCacheDefinitionN.bin` + `pivotCacheRecordsN.bin`
   built from `helpers.ts` `rec()` for each branch:
   - indexed-string field; numeric field; date field (incl. `hr != 0` to lock F4);
   - boolean; error; missing; `BrtPCRRecordDt` mode; `mdSxoper` 0x01/0x10 runs;
   - multiple caches; non-contiguous cache numbers (1 and 5).
2. A **§3.8 example fixture** (CustomerName / OrderDate / ProductName / UnitPrice / Quantity)
   — the only ground-truth-shaped case buildable without the real file.
3. Snapshot/assertion tests:
   - `fieldNames` + `rowCount` match current behaviour on these fixtures;
   - improved `rows`: dates as ISO, booleans as `true`/`false`, errors as `#ERR!` strings,
     numeric-as-text preserved as strings, non-Latin strings preserved.
4. Back-compat: existing `test/pivot-cache.test.ts` cases still pass (they use `rec(0x1b81…)`
   etc., which the new decoder also reads — no opcode change required).

## Out of scope (explicit)
- OLAP PivotCaches, MDX calculated members, slicer/timeline caches.
- `PCDISrvFmt` server formatting records (skipped, not required for value decoding).
- Changing the on-disk record opcodes (they are correct; see "Why" above).

## Rollout / versioning
- Ships as a minor/feature bump (non-breaking: `rows` element type widens to `PivotCacheCell`,
  still indexable; new `fields` array added).
- `CHANGELOG.md` notes: spec-driven decoder, streaming `iterPivotCacheRows`, fixed F2–F11,
  and that row values may differ (improved) vs older heuristic output.
