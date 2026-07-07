import js from '@eslint/js';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    rules: {
      // The codebase uses `any` for legacy error objects (e.g. `e: any` for
      // caught errors). Allow but discourage via warn.
      '@typescript-eslint/no-explicit-any': 'warn',
      // Allow unused function parameters prefixed with _.
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
        },
      ],
    },
  },
  {
    // Node scripts (server.mjs, eslint.config.js) — allow console/global.
    files: ['**/*.mjs'],
    languageOptions: {
      globals: {
        console: 'readonly',
        process: 'readonly',
        setTimeout: 'readonly',
        __dirname: 'readonly',
        require: 'readonly',
        module: 'readonly',
        exports: 'readonly',
      },
    },
  },
  {
    ignores: [
      'dist/',
      'node_modules/',
      'examples/browser-demo/bundle.js',
      'coverage/',
      'test/types.test-d.ts', // compile-time type assertions only
      'eslint.config.js',
    ],
  },
);
