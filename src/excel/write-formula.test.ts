import ExcelJS from 'exceljs';
import { expect, it } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { createWorkbook } from '../model/workbook.js';
import { exportXlsxWithReport } from './write.js';

it('SE-XLSX-2.1: omits unrepresentable cached text and retains the author’s formula', async () => {
  const formula = '=A2&A3&A4&A5';
  const workbook = createWorkbook([
    [formula, '=LEFT(A1,32767)', '="text"', '=1+'],
    ...Array.from({ length: 4 }, () => ['x'.repeat(8192)]),
  ]);
  const result = calculateWorkbook(workbook);
  expect(result.worksheets[0].values[0][0]).toBe('x'.repeat(32768));
  result.worksheets[0].values[0][2] = 'raw\u0001value';

  const { bytes, report } = await exportXlsxWithReport(workbook, result);
  const document = new ExcelJS.Workbook();
  await document.xlsx.load(bytes as never);
  const sheet = document.worksheets[0];
  expect(sheet.getCell('A1').formula).toBe(formula.slice(1));
  expect(sheet.getCell('A1').result).toBeUndefined();
  expect(sheet.getCell('B1').result).toBe('x'.repeat(32767));
  expect(sheet.getCell('C1').formula).toBe('"text"');
  expect(sheet.getCell('C1').result).toBeUndefined();
  expect(sheet.getCell('D1').value).toBe('=1+');
  expect(report.warnings).toEqual([
    expect.objectContaining({
      code: 'formula-cache',
      disposition: 'not-exported',
      count: 2,
      locations: ['Sheet1!A1', 'Sheet1!C1'],
    }),
    expect.objectContaining({ code: 'malformed-formula', locations: ['Sheet1!D1'] }),
  ]);
});

it('SE-XLSX-3.1: reports an unrepresentable rich result separately from formula syntax', async () => {
  const workbook = createWorkbook([['=RICH()', '=PLAIN()']]);
  const large = 'x'.repeat(32768);
  const result = calculateWorkbook(
    workbook,
    {},
    { callFunction: (name) => (name === 'RICH' ? `<p>${large}</p>` : large) },
  );

  const { bytes, report } = await exportXlsxWithReport(workbook, result, {
    formulaCapabilities: {
      hostFunctions: [{ name: 'RICH', capturedText: 'html-to-text' }, { name: 'PLAIN' }],
    },
  });
  const document = new ExcelJS.Workbook();
  await document.xlsx.load(bytes as never);
  expect(document.worksheets[0].getCell('A1').value).toBe('=RICH()');
  expect(document.worksheets[0].getCell('B1').value).toBe('=PLAIN()');
  expect(report.warnings).toEqual([
    expect.objectContaining({
      code: 'rich-formula-value',
      disposition: 'not-exported',
      locations: ['Sheet1!A1'],
    }),
    expect.objectContaining({
      code: 'captured-text',
      disposition: 'not-exported',
      locations: ['Sheet1!B1'],
    }),
    expect.objectContaining({
      code: 'formula-stored-result',
      disposition: 'not-exported',
      locations: ['Sheet1!B1'],
    }),
  ]);
});

it('SE-XLSX-2.1: rejects unrepresentable ordinary cell text with the cell location', async () => {
  const workbook = createWorkbook([['bad\u0001text']]);
  await expect(exportXlsxWithReport(workbook, calculateWorkbook(workbook))).rejects.toThrow(
    'Sheet1!A1: Text contains characters Excel cannot represent.',
  );
});

it('SE-XLSX-4.1: applies Excel precision and magnitude limits to stored numbers and formula caches', async () => {
  const workbook = createWorkbook([
    [
      '123456789012345',
      '1234567890123456',
      '1e-307',
      '5e-324',
      '1e308',
      '=1234567890123456',
      '=1e308',
      '=5e-324',
      '=1/3',
    ],
  ]);
  const { bytes, report } = await exportXlsxWithReport(workbook, calculateWorkbook(workbook));
  const document = new ExcelJS.Workbook();
  await document.xlsx.load(bytes as never);
  const sheet = document.worksheets[0];
  expect(sheet.getCell('A1').value).toBe(123456789012345);
  expect(sheet.getCell('B1').value).toBe('1234567890123456');
  expect(sheet.getCell('C1').value).toBe(1e-307);
  expect(sheet.getCell('D1').value).toBe('5e-324');
  expect(sheet.getCell('E1').value).toBe('1' + '0'.repeat(308));
  expect(sheet.getCell('F1').value).toMatchObject({
    formula: '1234567890123456',
    result: '1234567890123456',
  });
  expect(sheet.getCell('G1').value).toMatchObject({ formula: '1e308', result: 1e308 });
  expect(sheet.getCell('H1').value).toMatchObject({ formula: '5e-324', result: '5e-324' });
  expect(sheet.getCell('I1').value).toMatchObject({ formula: '1/3', result: '1/3' });
  expect(report.warnings).toEqual([
    expect.objectContaining({ code: 'numeric-text', disposition: 'stored-value', count: 6 }),
  ]);
});
