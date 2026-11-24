import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { createWorkbook } from '../model/workbook.js';
import { convertXlsx } from './convert.js';
import { exportXlsxWithReport } from './write.js';

describe('SE-XLSX-1.1: authored cell kinds and sparse geometry', () => {
  it('keeps exact numeric XML after a self-closing styled blank and imports blank row geometry', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Details');
    sheet.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF0000' } };
    sheet.getCell('B1').value = 0;
    sheet.getRow(5).height = 30;
    sheet.getRow(5).hidden = true;
    const files = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()));
    const path = 'xl/worksheets/sheet1.xml';
    const xml = strFromU8(files[path]);
    expect(xml).toContain('<c r="A1" s="1"/>');
    const decimal = '0.9999999999999999999999999';
    files[path] = strToU8(xml.replace('<v>0</v>', `<v>${decimal}</v>`));

    const converted = await convertXlsx(zipSync(files));
    const actual = converted.workbook.worksheets[0];
    expect(actual.data).toEqual([[null, decimal]]);
    expect(actual.style?.A1).toContain('background-color: #ff0000');
    expect(actual.cellProvenance?.A1).toBeUndefined();
    expect(actual.rows?.['4']).toEqual({ height: 40 });
    expect(converted.report.warnings).toContainEqual(
      expect.objectContaining({ code: 'hidden-row' }),
    );
  });

  it('retains text that resembles an authored number, a formula, an apostrophe escape or a quantity', async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Types');
    sheet.getRow(1).values = ['00123', '=1+1', "'as written", '3 kg'];
    sheet.getCell('D1').numFmt = '0.00 "kg"';
    sheet.getCell('E1').value = { error: '#DIV/0!' };
    sheet.getCell('F1').value = { formula: '1+1', result: 99 };

    const imported = await convertXlsx(new Uint8Array(await workbook.xlsx.writeBuffer()));
    const output = calculateWorkbook(imported.workbook);
    expect(output.worksheets[0].values[0].slice(0, 5)).toEqual([
      '00123',
      '=1+1',
      "'as written",
      '3 kg',
      '#DIV/0!',
    ]);
    expect(output.worksheets[0].values[0][5]).toMatchObject({ numerator: '2', denominator: '1' });
    expect(imported.report.formulas).toBe(1);
    expect(imported.report.warnings).toContainEqual(
      expect.objectContaining({ code: 'static-error' }),
    );

    const exported = await exportXlsxWithReport(imported.workbook, output);
    const opened = new ExcelJS.Workbook();
    await opened.xlsx.load(exported.bytes as never);
    expect(opened.worksheets[0].getCell('A1').value).toBe('00123');
    expect(opened.worksheets[0].getCell('B1').value).toBe('=1+1');
    expect(opened.worksheets[0].getCell('D1').value).toBe('3 kg');
    expect(opened.worksheets[0].getCell('F1').formula).toBe('1+1');
  });

  it('identifies the location when XLSX text exceeds the workbook text limit', async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet('Details').getCell('B2').value = 'x'.repeat(10_001);
    const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());
    await expect(convertXlsx(bytes)).rejects.toThrow(
      'Details!B2: cell exceeds the workbook text limit',
    );
  });
});

describe('SE-XLSX-4.1: number masks and sparse formatting', () => {
  it('separates masks from quoted text and reports unsupported modifiers and elapsed time', async () => {
    const workbook = new ExcelJS.Workbook();
    workbook.properties.date1904 = true;
    const sheet = workbook.addWorksheet('Formats');
    const masks = ['[Red]0.00', '0.0 "days"', '[h]:mm', 'mm', 'hh:mm:ss'];
    masks.forEach((mask, x) => {
      const cell = sheet.getCell(1, x + 1);
      cell.value = 0.5;
      cell.numFmt = mask;
    });

    const imported = await convertXlsx(new Uint8Array(await workbook.xlsx.writeBuffer()));
    const actual = imported.workbook.worksheets[0];
    expect(actual.data[0]).toEqual(['0.5', '0.5', '0.5', '1462.5', '0.5']);
    expect(actual.formats).toEqual({
      A1: { kind: 'number', decimals: 2 },
      B1: { kind: 'number', decimals: 1 },
      D1: { kind: 'date' },
      E1: { kind: 'time' },
    });
    const warningCodes = imported.report.warnings.map(({ code }) => code);
    expect(warningCodes).toEqual(
      expect.arrayContaining(masks.slice(0, 3).map((mask) => `number-format:${mask}`)),
    );
  });

  it('exports a style and number format outside the authored data matrix', async () => {
    const workbook = createWorkbook([['base']]);
    const sheet = workbook.worksheets[0];
    sheet.style = { B5: 'background-color: #ff0000' };
    sheet.formats = { B5: { kind: 'currency', currency: 'EUR' } };
    const exported = await exportXlsxWithReport(workbook, calculateWorkbook(workbook));
    const opened = new ExcelJS.Workbook();
    await opened.xlsx.load(exported.bytes as never);
    const blank = opened.worksheets[0].getCell('B5');
    expect(blank.value).toBeNull();
    expect(blank.fill).toMatchObject({ fgColor: { argb: 'FFff0000' } });
    expect(blank.numFmt).toContain('[$EUR-409]');
    const imported = await convertXlsx(exported.bytes);
    expect(imported.workbook.worksheets[0].data).toEqual([['base']]);
    expect(imported.workbook.worksheets[0].style?.B5).toContain('background-color: #ff0000');
  });
});
