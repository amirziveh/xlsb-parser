# Contributing

Contributions are welcome. This project follows a TDD discipline.

## Setup

```sh
git clone https://github.com/anomalyco/xlsb-parser
cd xlsb-parser
npm ci
npm test
```

## Workflow

1. **Write a failing test first.** Watch it fail. Then implement the
   minimal code to make it pass. See `test/` for the existing style.
2. **Run the full verification suite before pushing:**

   ```sh
   npm test                # vitest
   npm run test:coverage   # enforce 90% statements / 100% functions
   npm run lint            # eslint
   npm run format:check    # biome
   npm run test:types      # tsd
   npx tsc --noEmit        # strict typecheck
   npm run build:browser    # rebuild browser bundle
   ```

3. **Commit messages**: follow the existing style — a short subject line
   prefixed with `feat:`, `fix:`, `refactor:`, `test:`, `docs:`, or
   `chore:`. Wrap the body at 72 columns and explain *why*, not just what.

4. **Don't commit `dist/`** — it's gitignored and built at publish time.
   The browser demo's `bundle.js` under `examples/browser-demo/` *is*
   committed; rebuild it with `npm run build:browser` if you change `src/`.

5. **Coverage gate**: the CI will fail if statement coverage on `src/`
   drops below 90% (excluding `pivot-cache.ts`, which is the heuristic
   decoder slated for a spec-driven rewrite).

## Adding a new cell type or record opcode

1. Add the opcode constant to `src/record-stream.ts`.
2. Add a case to `readCell` / `readShortCell` in `src/sheet.ts`.
3. Add a builder to `test/helpers.ts` and a test in `test/cell-types.test.ts`.
4. Watch the test fail (verify RED), then implement, then watch green.

## Releases

Releases are cut by tagging `v*` and pushing — the GitHub Actions
`publish.yml` workflow builds, runs tests, and publishes to npm with
`--provenance`. Follow [Keep a Changelog](https://keepachangelog.com/)
format in `CHANGELOG.md`.
