import assert from 'node:assert/strict';
import { access, readFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const manifest = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
for (const entry of Object.values(manifest.exports)) {
  for (const file of typeof entry === 'string' ? [entry] : Object.values(entry))
    await access(new URL(`../${file}`, import.meta.url));
}
// These entry points must remain importable in a plain Node process with no DOM globals.
assert.equal(typeof globalThis.window, 'undefined');
assert.equal(typeof globalThis.document, 'undefined');
const { createWorkbook, assertWorkbook } =
  await import('@flamboyantengineer/spreadsheet-editor/model');
const {
  calculateWorkbook,
  workbookCalculationKey,
  assertExternalCalculationSnapshot,
  calculationCellKey,
} = await import('@flamboyantengineer/spreadsheet-editor/calculation');
const { createDocumentInsertion } = await import('@flamboyantengineer/spreadsheet-editor/document');
const { createMemoryHost, resolveSpreadsheetMode } =
  await import('@flamboyantengineer/spreadsheet-editor/host');
assert.equal(
  resolveSpreadsheetMode(await createMemoryHost().permissions.capabilities('a')),
  'author',
);
const { SpreadsheetEditor } = await import('@flamboyantengineer/spreadsheet-editor');
assert.equal(typeof SpreadsheetEditor, 'function');
const workbook = createWorkbook([['=1-(1-1E-25)']]);
const result = calculateWorkbook(workbook);
assert.equal(result.worksheets[0].values[0][0].denominator, '10000000000000000000000000');
assert.equal(createDocumentInsertion(workbook, result, { mode: 'table' }).kind, 'table');
console.log('Built exports and headless Node imports passed.');

const canonicalWorkbook = createWorkbook([['authored text in another input unit']]);
canonicalWorkbook.worksheets[0].inputValues = {
  A1: {
    kind: 'number',
    numerator: '1',
    denominator: '10000000000000000000000000',
    approximate: false,
    provenance: { constants: 'revision-1' },
  },
};
assertWorkbook(canonicalWorkbook);
const canonicalResult = calculateWorkbook(canonicalWorkbook);
assert.equal(canonicalResult.worksheets[0].values[0][0].denominator, '10000000000000000000000000');
const inputKey = workbookCalculationKey(canonicalWorkbook);
const snapshot = {
  inputKey,
  revision: 1,
  status: 'ready',
  result: canonicalResult,
  cellStates: { [calculationCellKey('sheet-1', 0, 0)]: 'ready' },
  hostContext: { revision: 'captured' },
};
assertExternalCalculationSnapshot(snapshot, canonicalWorkbook);
assert.throws(
  () => assertExternalCalculationSnapshot({ ...snapshot, result: undefined }, canonicalWorkbook),
  /result/i,
);
assert.throws(
  () => assertExternalCalculationSnapshot({ ...snapshot, inputKey: 'obsolete' }, canonicalWorkbook),
  /key/i,
);
assert.throws(
  () =>
    assertExternalCalculationSnapshot(
      { ...snapshot, status: 'pending', cellStates: undefined },
      canonicalWorkbook,
    ),
  /cached/i,
);
const visuallyChanged = structuredClone(canonicalWorkbook);
visuallyChanged.name = 'A different title';
visuallyChanged.worksheets[0].style = { A1: 'font-weight: bold' };
assert.equal(workbookCalculationKey(visuallyChanged), inputKey);
console.log('Built canonical input and external-publication contracts passed.');

// Explicit files + NodeNext force a consumer to resolve published .d.ts exports, not src/.
execFileSync(
  process.execPath,
  [
    fileURLToPath(new URL('../node_modules/typescript/bin/tsc', import.meta.url)),
    '--ignoreConfig',
    '--noEmit',
    '--strict',
    '--skipLibCheck',
    '--target',
    'ES2022',
    '--module',
    'NodeNext',
    '--moduleResolution',
    'NodeNext',
    '--jsx',
    'react-jsx',
    '--lib',
    'ES2023,DOM',
    '--types',
    'node',
    fileURLToPath(new URL('./fixtures/public-host.tsx', import.meta.url)),
    fileURLToPath(new URL('../examples/host-integration.tsx', import.meta.url)),
  ],
  { cwd: fileURLToPath(new URL('../', import.meta.url)), stdio: 'inherit' },
);
console.log('Built public host types, example and legacy consumers passed.');

const { importXlsx, exportXlsx } = await import('@flamboyantengineer/spreadsheet-editor/excel');
const xlsx = await exportXlsx(workbook, result);
const imported = await importXlsx(xlsx);
assert.equal(imported.workbook.worksheets[0].data[0][0], '=1-(1-1E-25)');
assert.equal(imported.report.formulas, 1);
await assert.rejects(importXlsx(new Uint8Array([0, 1, 2, 3])));
console.log('Isolated Excel import and stored-result export passed.');
