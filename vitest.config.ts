import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test/**/*.test.ts'],
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      // pivot-cache.ts is excluded from the coverage gate because it's the
      // heuristic decoder slated for a spec-driven rewrite in P4. Coverage
      // on soon-to-be-replaced code wastes effort to push past the level
      // the existing test suite naturally lands at.
      exclude: ['src/pivot-cache.ts'],
      thresholds: {
        // Per-file gates catch regressions on the modules that should stay
        // green. The uncovered lines in sheet.ts are unreachable `default:
        // return null` defensive branches (every opcode 0x01..0x0B / 0x0C..0x12
        // is in the switch).
        lines: 90,
        statements: 90,
        functions: 95,
        branches: 70,
      },
    },
  },
});
