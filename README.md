# xlsb-parser

Browser-based **XLSB** (Excel Binary Workbook) parser. Pure JS, no native addons. Works in the browser and Node.js 18+.

## Install

```sh
npm install xlsb-parser
```

## Usage

```ts
import { parseXlsb } from 'xlsb-parser';

const response = await fetch('report.xlsb');
const buffer = await response.arrayBuffer();
const wb = await parseXlsb(buffer);

console.log(wb.sheets[0].name);
// → 'Sheet1'
console.log(wb.sheets[0].rows[0].cols[0]);
// → { t: 's', v: 'Hello' }
```

### API

**`parseXlsb(data: ArrayBuffer | Uint8Array, onProgress?: (msg, pct) => void): Promise<ParsedXlsb>`**

| Field | Type | Description |
|-------|------|-------------|
| `sheets` | `Sheet[]` | Parsed worksheets |
| `sharedStrings` | `string[]` | All shared strings |
| `pivotCaches` | `PivotCacheTable[]` | Pivot cache data |
| `binaryDumps` | `BinaryDump[]` | Raw record dump per `.bin` file |
| `xmlFiles` | `Record<string, string>` | Raw XML content |
| `summary` | `{ fileCount, totalRecords }` | Counts |

**`Sheet`**: `{ name: string, rows: ParsedRow[], totalCells: number }`

**`ParsedRow`**: `{ row: number, cols: Record<number, Cell> }` — `cols` key is the zero-based column index.

**`Cell`**:

| `t` | Meaning | `v` type |
|-----|---------|----------|
| `n` | Number | `number` |
| `s` | String | `string` |
| `b` | Boolean | `boolean` |
| `e` | Error | `string` (the error code) |
| `blank` | Empty | `undefined` |

### Progress callback

```ts
const wb = await parseXlsb(buffer, (msg, pct) => {
  console.log(`[${pct}%] ${msg}`);
});
```

## Browser

Include via `<script type="module">`:

```html
<script type="module">
import { parseXlsb } from './node_modules/xlsb-parser/dist/index.js';
// or from a CDN
</script>
```

Or use the pre-built bundle at `public/bundle.js`.

## CLI

Quick inspection from Node:

```sh
node -e "import { parseXlsb } from 'xlsb-parser'; import { readFileSync } from 'fs';
const w = await parseXlsb(readFileSync(process.argv[1]));
w.sheets.forEach(s => console.log(s.name, s.rows.length + ' rows'));" file.xlsb
```

## Commands

```sh
npm install      # install dependencies
npm run build    # compile TypeScript → dist/
npm run dev      # build + serve browser demo on localhost:8080
npm test         # run tests (vitest)
```

## License

GPL-3.0
