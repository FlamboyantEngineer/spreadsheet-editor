import js from '@eslint/js';
import globals from 'globals';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  {
    ignores: [
      'node_modules/**',
      'dist/**',
      'labs/build/**',
      'work/**',
      'test-results/**',
      'playwright-report/**',
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  { languageOptions: { globals: { ...globals.browser, ...globals.node } } },
  {
    files: ['src/**/*.{ts,tsx}'],
    rules: {
      'no-restricted-globals': ['error', 'localStorage', 'sessionStorage'],
      'no-restricted-syntax': [
        'error',
        {
          selector: 'CallExpression[callee.name="eval"]',
          message: 'Formulas must use the parsed syntax tree.',
        },
        {
          selector: 'NewExpression[callee.name="Function"]',
          message: 'Formulas must use the parsed syntax tree.',
        },
      ],
    },
  },
);
