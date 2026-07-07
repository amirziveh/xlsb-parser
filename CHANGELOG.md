# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- **Spec-driven pivot cache decoder** (opt-in `parsePivotCaches: true`, replaces the
  heuristic sampler). The new decoder reads authoritative `BrtBeginPCDField` field types,
  `BrtBeginPCDFAtbl` flags (`fNumField`, `fDateInField`, `fHasTextItem`), and
  `BrtBeginPCDIRun`/`BrtPCDI*` shared-item records from the definition part, producing
  per-field `PivotCacheField` descriptors with correct `kind` and `sharedItems`.
- `PivotCacheCell` discriminated union (`'s' | 'n' | 'd' | 'b' | 'e' | 'blank'`) for
  decoded cache cell values. Rows remain indexable as `PivotCacheCell[][]`.
- `PivotCacheField` and `PivotCacheSummary` types.
- **Streaming pivot cache rows** via `openXlsb(..., { parsePivotCaches: true })`:
  cache **definitions** are parsed eagerly, **rows** stream lazily through
  `handle.iterPivotCacheRows(indexOrName, { maxRows?, onProgress? })`.
  `handle.collectPivotCache(indexOrName)` drains the full cache.
- `parsePivotCache` is guarded within `parseXlsb`/`openXlsb` — one malformed
  cache no longer aborts the entire workbook (F11).

### Fixed
- **PCDIDateTime** day-of-month now read as `u8` (was `u16`, corrupting datetimes
  with non-zero hours; F4).
- **BrtBeginPCDIRun** now handles `mdSxoper` 0x01 (number runs) and 0x10 (error
  runs) in addition to 0x02/0x20 (F5).
- **BrtPCRRecordDt** (per-field-value records) rows now decode instead of being
  silently skipped (F2).
- **String fields** no longer require `[A-Za-z]` characters — numeric-as-text
  and non-Latin text are preserved (F8).
- **Field names** are taken from the definition part rather than truncated by
  the first 5 data rows (F10).
- `BRT_PCDI_STRING2` (0x001f) string shared items are handled alongside 0x0018.

## [1.0.0-rc.1] — 2026-07-07

### Added
- **Options bag** for `parseXlsb`: `{ onProgress?, maxZipBytes?, maxRowsPerSheet?, dumpBinaries?, readXml?, parsePivotCaches? }`.
- **Streaming API**: `openXlsb(data, options?)` returns an `XlsbHandle` whose `iterSheetRows(sheetIndex)` is an async generator yielding rows lazily — O(cells-per-row) memory instead of O(N).
- **Styles + date detection**: parses `xl/styles.bin`, exposes `cell.numFmtId`, `cell.isDate`, `cell.dateValue` (ISO 8601). The raw serial value `v` stays unchanged.
- **`XlsbSizeError`** (subclass of `Error` with `.limit` / `.actual`) raised when `maxZipBytes` or `maxRowsPerSheet` caps are exceeded.
- **Pivot caches now opt-in** via `parsePivotCaches: true` (default `false`). The decoder remains heuristic for opt-in callers; a spec-driven rewrite is planned for a future release.
- **Type tests** via `tsd` (`test/types.test-d.ts`).
- **Linting** via `eslint` + `@typescript-eslint`, **formatting** via `biome`.
- **CI**: macos-latest added to the test matrix; `npm ci` for reproducibility; lint + format-check + type-test job; npm publish runs with `--provenance`.
- `CHANGELOG.md`, `SECURITY.md`, `CONTRIBUTING.md`.
- `README.md` rewritten with the new API and a "Limitations" section.

### Changed
- `parseXlsb` signature is now `(data, options?)`. A function passed as the 2nd arg is still treated as `onProgress` for backwards compatibility through 1.x; this legacy form will be removed at 2.0.
- `dumpBinaries`, `readXml` default to `false` (were always-on in 0.2.0). Use the options bag to opt in.
- Browser demo moved from `public/` to `examples/browser-demo/`.
- `tsconfig.json` tightened: `noUnusedLocals`, `noUnusedParameters`, `noFallthroughCasesInSwitch`, `noImplicitOverride`, `sourceMap`, `declarationMap`.
- Minimum Node version is now 20 (Node 18 is EOL).

### Fixed
- **iStyleRef sign extension**: previously checked bit 7 of the low byte (`d[4]&0x80`); now correctly checks bit 15 of the high byte (`d[5]&0x80`). Style indices in the range 128–255 were silently corrupted.
- **Bounds checks in `readCell` / `readShortCell`**: every case now validates `d.length >= off + valueSize` and returns `null` when insufficient bytes remain. Previously, truncated cell records silently read adjacent buffer memory.
- **`records()` throws on truncation** instead of silently clamping the declared record size to whatever bytes remain. The error names the record's offset and type.
- **`BRT_FMLA_STRING` length guard**: `off + 6` → `off + 4` (the minimum for `cch`). Cells with exactly 4 or 5 bytes of payload no longer return `''`.
- **RK decode 67× faster**: replaced per-call `BigInt` arithmetic + throwaway `ArrayBuffer` with a pre-allocated module-scope `DataView` / `Uint32Array`. Benchmarked at 6537 ms → 98 ms over 10M decodes.
- **Module split**: `src/index.ts` (689-line monolith) → 9 focused modules: `types.ts`, `record-stream.ts`, `workbook.ts`, `shared-strings.ts`, `sheet.ts`, `styles.ts`, `pivot-cache.ts`, `dump.ts`, `handle.ts`, slim `index.ts` orchestrator.

### License
- Switched from **GPL-3.0** to **MIT** to match the rest of the npm spreadsheet-ecosystem.

## [0.2.0] — 2026-07-02

Initial public prototype.
