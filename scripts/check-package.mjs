import assert from 'node:assert/strict';
import { access, readFile } from 'node:fs/promises';

const manifest = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));
for (const entry of Object.values(manifest.exports)) {
  for (const file of typeof entry === 'string' ? [entry] : Object.values(entry))
    await access(new URL(`../${file}`, import.meta.url));
}
// These entry points must remain importable in a plain Node process with no DOM globals.
const { createWorkbook } = await import('@flamboyantengineer/spreadsheet-editor/model');
const { calculateWorkbook } = await import('@flamboyantengineer/spreadsheet-editor/calculation');
const { createDocumentInsertion } = await import('@flamboyantengineer/spreadsheet-editor/document');
const { SpreadsheetEditor } = await import('@flamboyantengineer/spreadsheet-editor');
assert.equal(typeof SpreadsheetEditor, 'function');
const workbook = createWorkbook([['=1-(1-1E-25)']]);
const result = calculateWorkbook(workbook);
assert.equal(result.worksheets[0].values[0][0].denominator, '10000000000000000000000000');
assert.equal(createDocumentInsertion(workbook, result, { mode: 'table' }).kind, 'table');
console.log('Built exports and headless Node imports passed.');

const { importXlsx, exportXlsx } = await import('@flamboyantengineer/spreadsheet-editor/excel');
const xlsx = await exportXlsx(workbook, result);
const imported = await importXlsx(xlsx);
assert.equal(imported.workbook.worksheets[0].data[0][0], '=1-(1-1E-25)');
assert.equal(imported.report.formulas, 1);
await assert.rejects(importXlsx(new Uint8Array([0, 1, 2, 3])));
console.log('Isolated Excel import and stored-result export passed.');
