import { createBaseConfig, ENUM_SYNTAX } from '@flamboyantengineer/ui-components/tooling/eslint';
import reactHooks from 'eslint-plugin-react-hooks';
import tseslint from 'typescript-eslint';

// The React Compiler's rule set, every rule at error: code the compiler cannot optimise fails lint.
const REACT_COMPILER_RULES = Object.fromEntries(
  Object.keys(reactHooks.configs['recommended-latest'].rules).map((rule) => [rule, 'error']),
);

const TYPED_FILES = ['src/**/*.{ts,tsx}', 'labs/src/**/*.{ts,tsx}', 'e2e/**/*.ts', '*.config.ts'];

// Formulas evaluate through the parsed syntax tree, never through generated code.
const FORMULA_SYNTAX = [
  {
    selector: 'CallExpression[callee.name="eval"]',
    message: 'Formulas must use the parsed syntax tree.',
  },
  {
    selector: 'NewExpression[callee.name="Function"]',
    message: 'Formulas must use the parsed syntax tree.',
  },
];

const SIZE_LIMITS_PENDING = [
  'src/calculation/**',
  'src/document/**',
  'src/editor/**',
  'src/engine/**',
  'src/excel/**',
  'src/host/**',
  'src/model/**',
];

const STORAGE_GLOBALS = ['localStorage', 'sessionStorage', 'indexedDB'].map((name) => ({
  name,
  message: 'Persist through a host callback (specs/HOST-INTEGRATION.md).',
}));

export default tseslint.config(
  // The shared rules (ui-components' specs/TOOLING.md): imports, the type-aware TypeScript rules,
  // named exports, and the size and naming limits of library source.
  ...createBaseConfig({
    name: 'spreadsheet-editor',
    tsconfigRootDir: import.meta.dirname,
    typedFiles: TYPED_FILES,
    // Vitest and Playwright load their configuration through its default export.
    defaultExportLoaders: ['*.config.ts'],
    sourceFiles: ['src/**/*.{ts,tsx}'],
    nodeFiles: ['scripts/**/*.mjs', '*.config.{js,mjs,ts}', 'eslint.config.js'],
    // The example is type-checked against the built package by `npm run check:package`.
    ignores: ['labs/build/', 'playwright-report/', 'work/', 'examples/'],
    components: true,
  }),
  {
    files: TYPED_FILES,
    rules: { 'no-restricted-syntax': ['error', ENUM_SYNTAX] },
  },
  {
    files: ['src/**/*.{ts,tsx}', 'labs/src/**/*.{ts,tsx}'],
    plugins: { 'react-hooks': reactHooks },
    rules: { ...REACT_COMPILER_RULES },
  },
  {
    files: ['src/**/*.{ts,tsx}'],
    ignores: ['src/**/*.test.{ts,tsx}'],
    rules: {
      'no-restricted-globals': ['error', ...STORAGE_GLOBALS],
      'no-restricted-syntax': ['error', ENUM_SYNTAX, ...FORMULA_SYNTAX],
    },
  },
  {
    // The shared size limits are being adopted area by area: each refactor pull request removes its
    // area from this list until it is empty.
    files: SIZE_LIMITS_PENDING,
    ignores: ['src/**/*.test.{ts,tsx}'],
    rules: {
      complexity: 'off',
      'max-depth': 'off',
      'max-params': 'off',
      'max-lines': 'off',
      'max-lines-per-function': 'off',
    },
  },
);
