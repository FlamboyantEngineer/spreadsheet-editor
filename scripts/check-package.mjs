import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { access, readFile } from 'node:fs/promises';
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
const { createMemoryHost, resolveSpreadsheetMode, createTransactionReceiptTracker } =
  await import('@flamboyantengineer/spreadsheet-editor/host');
assert.equal(
  resolveSpreadsheetMode(await createMemoryHost().permissions.capabilities('a')),
  'author',
);
const receipts = createTransactionReceiptTracker();
const acceptedTransaction = { transactionId: 'existing-editor-transaction' };
receipts.record(acceptedTransaction.transactionId, 1);
assert.equal(receipts.acknowledge(acceptedTransaction.transactionId), 1);
assert.equal(
  'createTransactionReceiptTracker' in
    (await import('@flamboyantengineer/spreadsheet-editor/model')),
  false,
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

// Exercise the actual tarball as a downstream package, including files/exports and CSS.
const { mkdtemp, mkdir, writeFile, symlink, cp, rm } = await import('node:fs/promises');
const { tmpdir } = await import('node:os');
const { join } = await import('node:path');
const consumer = await mkdtemp(join(tmpdir(), 'spreadsheet-consumer-'));
try {
  const root = fileURLToPath(new URL('../', import.meta.url));
  const packed = JSON.parse(
    execFileSync(
      process.env.npm_execpath ? process.execPath : 'npm',
      [
        ...(process.env.npm_execpath ? [process.env.npm_execpath] : []),
        'pack',
        '--ignore-scripts',
        '--json',
        '--pack-destination',
        consumer,
      ],
      {
        cwd: root,
        encoding: 'utf8',
      },
    ),
  );
  const unpacked = join(consumer, 'unpacked');
  await mkdir(unpacked);
  execFileSync('tar', ['-xzf', join(consumer, packed[0].filename), '-C', unpacked]);
  const packageRoot = join(unpacked, 'package');
  await symlink(join(root, 'node_modules'), join(packageRoot, 'node_modules'), 'junction');
  await mkdir(join(consumer, 'node_modules', '@flamboyantengineer'), { recursive: true });
  await symlink(
    packageRoot,
    join(consumer, 'node_modules', '@flamboyantengineer', 'spreadsheet-editor'),
    'junction',
  );
  for (const name of ['react', 'react-dom', '@types'])
    await symlink(
      join(root, 'node_modules', name),
      join(consumer, 'node_modules', name),
      'junction',
    );
  await writeFile(
    join(consumer, 'package.json'),
    JSON.stringify({ type: 'module', private: true }),
  );
  await cp(new URL('./fixtures/public-host.tsx', import.meta.url), join(consumer, 'consumer.tsx'));
  await cp(
    new URL('../examples/host-integration.tsx', import.meta.url),
    join(consumer, 'host.tsx'),
  );
  await writeFile(
    join(consumer, 'consumer.mjs'),
    `
    import assert from 'node:assert/strict';
    import { readFile } from 'node:fs/promises';
    import { createWorkbook, createWorkbookOperationBinding } from '@flamboyantengineer/spreadsheet-editor/model';
    import * as model from '@flamboyantengineer/spreadsheet-editor/model';
    import { createTransactionReceiptTracker } from '@flamboyantengineer/spreadsheet-editor/host';
    import { createLinkedSpreadsheetInsertion } from '@flamboyantengineer/spreadsheet-editor/document';
    import { SpreadsheetEditor } from '@flamboyantengineer/spreadsheet-editor';
    assert.equal(createWorkbook().schemaVersion, 1);
    assert.equal(typeof createWorkbookOperationBinding, 'function');
    assert.equal('createTransactionReceiptTracker' in model, false);
    const receipts = createTransactionReceiptTracker();
    const accepted = { transactionId: 'existing-editor-transaction' };
    receipts.record(accepted.transactionId, 1);
    assert.equal(receipts.acknowledge(accepted.transactionId), 1);
    assert.equal(typeof createLinkedSpreadsheetInsertion, 'function');
    assert.equal(typeof SpreadsheetEditor, 'function');
    const referenceUrl = new URL(import.meta.resolve('@flamboyantengineer/spreadsheet-editor/reference.css'));
    const css = await readFile(referenceUrl, 'utf8');
    assert.ok(css.includes('@flamboyantengineer/ui-components/reference.css'));
    assert.ok(css.includes("@import './editor.css'"));
    assert.ok((await readFile(new URL('editor.css', referenceUrl), 'utf8')).includes('.se-editor'));
    const shadcn = await readFile(new URL(import.meta.resolve('@flamboyantengineer/spreadsheet-editor/shadcn.css')), 'utf8');
    assert.ok(shadcn.includes('@flamboyantengineer/ui-components/shadcn.css'));
    assert.ok(shadcn.includes("@import './editor.css'"));
  `,
  );
  execFileSync(process.execPath, [join(consumer, 'consumer.mjs')], {
    cwd: consumer,
    stdio: 'inherit',
  });
  execFileSync(
    process.execPath,
    [
      join(root, 'node_modules/typescript/bin/tsc'),
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
      join(consumer, 'consumer.tsx'),
      join(consumer, 'host.tsx'),
    ],
    { cwd: consumer, stdio: 'inherit' },
  );
  const { build } = await import('esbuild');
  for (const [sheet, marker] of [
    ['reference.css', '--ref-'],
    ['shadcn.css', '--primary-foreground'],
  ]) {
    await writeFile(
      join(consumer, sheet),
      `@import '@flamboyantengineer/spreadsheet-editor/${sheet}';`,
    );
    const stylesheet = await build({
      absWorkingDir: consumer,
      entryPoints: [sheet],
      bundle: true,
      write: false,
    });
    assert.ok(stylesheet.outputFiles[0].text.includes('.se-editor'), sheet);
    assert.ok(stylesheet.outputFiles[0].text.includes(marker), sheet);
  }
  console.log('Packed public consumer imports, declarations and opt-in stylesheets passed.');
} finally {
  await rm(consumer, { recursive: true, force: true });
}
