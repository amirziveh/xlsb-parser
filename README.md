# xlsb-parser

Pure-JS **XLSB** (Excel Binary Workbook) parser. No native addons. Works in
browsers and Node.js 20+.

## Install

```sh
npm install xlsb-parser
```

## Quick start

```ts
import { parseXlsb } from 'xlsb-parser';

const buffer = await file.arrayBuffer();
const wb = await parseXlsb(buffer);

console.log(wb.sheets[0].name);                 // → 'Sheet1'
console.log(wb.sheets[0].rows[0].cols[0]);      // → { t: 's', v: 'Hello' }
```

## API

### `parseXlsb(data, options?)`

```ts
function parseXlsb(
  data: ArrayBuffer | Uint8Array,
  options?: ParseOptions,
): Promise<ParsedXlsb>;

interface ParseOptions {
  onProgress?: (msg: string, pct: number) => void;
  maxZipBytes?: number;          // throw XlsbSizeError if decompressed > N
  maxRowsPerSheet?: number;      // stop a sheet after N rows
  dumpBinaries?: boolean;         // default false. Populate binaryDumps.
  readXml?: boolean;              // default false. Populate xmlFiles.
  parsePivotCaches?: boolean;    // default false. Spec-driven pivot cache decoder.
}
```

**Backwards compat (deprecated 1.x only):** passing a function as the 2nd
arg is treated as `onProgress`. This form will be removed at 2.0.

Returns `ParsedXlsb`:

| Field | Type | Description |
|-------|------|-------------|
| `sheets` | `Sheet[]` | Worksheets (parsed lazily when using `openXlsb`) |
| `sharedStrings` | `string[]` | All shared strings |
| `styles` | `StylesTable \| null` | Cell style table (since 1.0) |
| `pivotCaches` | `PivotCacheTable[]` | Pivot caches (opt-in via `parsePivotCaches: true`). Each entry has `.name`, `.fieldNames` (deprecated), `.rows`, `.fields` (`PivotCacheField[]` with correct `kind`/`sharedItems`), `.summary`. |
| `binaryDumps` | `BinaryDump[]` | Debug record dumps (opt-in via `dumpBinaries: true`) |
| `xmlFiles` | `Record<string, string>` | Raw XML/rels content (opt-in via `readXml: true`) |
| `summary` | `{ fileCount, totalRecords }` | Counts |

### `openXlsb(data, options?)` — streaming handle

For huge sheets, use the streaming API to iterate rows lazily without
buffering the whole sheet:

```ts
import { openXlsb } from 'xlsb-parser';

const handle = await openXlsb(buffer);
console.log(handle.sheetNames);             // ['Sheet1', 'Sheet2']

for await (const row of handle.iterSheetRows(0)) {
  // Process one row at a time. Memory: O(cells per row), not O(total rows).
  console.log(row.row, Object.keys(row.cols).length);
  if (row.row > 1000) break;                  // stop early if you want
}

// Or drain an entire sheet (still streams internally):
const sheet = await handle.collectSheet(0);
```

`iterSheetRows(index, { maxRows?, onProgress? })` accepts an optional cap.

When `parsePivotCaches: true` was passed at open time, the handle also exposes pivot cache metadata and streaming:

- `handle.pivotCaches` — `PivotCacheSummary[]` with eager field definitions.
- `handle.iterPivotCacheRows(indexOrName, { maxRows?, onProgress? })` — async
  generator yielding `PivotCacheCell[]` rows, O(cells-per-row) memory.
- `handle.collectPivotCache(indexOrName)` — drain the full cache into a
  `PivotCacheTable`.

### Cell type

```ts
interface Cell {
  t: 'n' | 's' | 'b' | 'e' | 'blank' | 'f';
  v?: number | string | boolean;
  err?: string;                  // present when t === 'e'
  ixf?: number;                  // signed iStyleRef into the cellXfs table
  numFmtId?: number;             // numeric format ID (since 1.0)
  isDate?: boolean;              // true when numFmtId is a date/time format
  dateValue?: string;            // ISO 8601 string when isDate && t === 'n'
}
```

| `t` | Meaning | `v` type |
|-----|---------|----------|
| `n` | Number | `number` |
| `s` | String | `string` |
| `b` | Boolean | `boolean` |
| `e` | Error | `err: string` |
| `blank` | Empty | `undefined` |

When `xl/styles.bin` is present, numeric cells whose style is a date format
get `numFmtId`, `isDate: true`, and `dateValue` (ISO 8601). The raw serial
number is preserved in `v` — no silent type coercion.

```ts
const cell = wb.sheets[0].rows[0].cols[0];
// cell = { t: 'n', v: 44927, numFmtId: 14, isDate: true, dateValue: '2023-01-01T00:00:00.000Z' }
```

### `XlsbSizeError`

Thrown when `maxZipBytes` or `maxRowsPerSheet` caps are exceeded. Subclass
of `Error` with `.limit` and `.actual` fields so callers can differentiate
"too big" from "broken".

## Browser

```html
<script type="module">
import { parseXlsb } from './node_modules/xlsb-parser/dist/index.js';
</script>
```

A pre-built demo bundle lives in `examples/browser-demo/`. Run it with
`npm run dev`.

## CLI (quick inspection)

```sh
node -e "import { parseXlsb } from 'xlsb-parser'; import { readFileSync } from 'fs';
const w = await parseXlsb(readFileSync(process.argv[2]));
w.sheets.forEach(s => console.log(s.name, s.rows.length + ' rows'));" file.xlsb
```

## Limitations

- **Pivot caches (opt-in):** now spec-driven (MS-XLSB §2.1.7.38/§2.1.7.39)
  for non-OLAP caches. OLAP/slicer/timeline/server-formatting caches remain
  best-effort.
- **No streaming unzip**: `openXlsb()` streams *rows*, but the ZIP step
  itself is buffered. For files >1 GB consider pre-decompressing. Use
  `maxZipBytes` to refuse oversized inputs.
- **Rich-text formatting runs**: cell strings decode correctly but
  formatting runs (italic, bold, color) are discarded — only the text is
  returned.
- **Date detection is heuristic**: built-in date format IDs (14–22, 27–36,
  45–47, 50–58, 78–81) plus custom format strings containing `d`/`m`/`y`/
  `h`/`s` outside escapes. Edge cases with unusual custom formats may be
  missed.

## Commands

```sh
npm install              # install dependencies
npm run build            # compile TypeScript → dist/
npm run build:browser    # rebuild the browser demo bundle
npm test                 # run tests (vitest)
npm run test:coverage    # tests + coverage gate (90% stmts, 100% funcs)
npm run test:types       # compile-time type assertions
npm run lint             # eslint
npm run format           # biome format (write)
npm run dev              # build browser bundle + serve demo on :8080
```

## License

MIT
