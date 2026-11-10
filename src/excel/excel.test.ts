import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { convertXlsx } from './convert.js';
import { inspectXlsx } from './guard.js';
import { exportXlsx } from './write.js';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { createWorkbook } from '../model/workbook.js';

describe('Host SE-XL-1/2: Excel conversion and loss reporting', () => {
  it('keeps formulas and geometry while discarding caches and reporting unsupported content', async () => {
    const document = new ExcelJS.Workbook(),
      sheet = document.addWorksheet('Inputs');
    sheet.getCell('A1').value = 'Item';
    sheet.getCell('B1').value = { formula: '1+1', result: 9999 };
    sheet.getCell('A2').value = { richText: [{ text: 'Rich ' }, { text: 'text' }] };
    sheet.getCell('B2').value = 12.5;
    sheet.getCell('B2').numFmt = '0.00%';
    sheet.getCell('B2').font = { bold: true, color: { argb: 'FF175CD3' } };
    sheet.getCell('A3').value = { formula: 'ACCRINT(1,2,3)', result: 99 };
    sheet.getCell('B3').value = 'ITAR controlled';
    sheet.getColumn(1).width = 20;
    sheet.getColumn(2).hidden = true;
    sheet.getRow(2).height = 24;
    sheet.mergeCells('A4:B4');
    sheet.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
    const result = await convertXlsx(new Uint8Array(await document.xlsx.writeBuffer()));
    const workbook = result.workbook,
      converted = workbook.worksheets[0];
    expect(converted.data[0][1]).toBe('=1+1');
    expect(calculateWorkbook(workbook).worksheets[0].values[0][1]).toMatchObject({
      numerator: '2',
    });
    expect(converted.data[1][0]).toBe('Rich text');
    expect(converted.formats!.B2).toMatchObject({ kind: 'percent', decimals: 2 });
    expect(converted.style!.B2).toContain('font-weight:bold');
    expect(converted.style!.B2).toContain('#175cd3');
    expect(converted.columns![0].width).toBe(145);
    expect(converted.rows!['1'].height).toBe(32);
    expect(converted.mergeCells!.A4).toEqual([2, 1]);
    expect(converted.freezeColumns).toBe(1);
    expect(result.report.warnings.map((warning) => warning.code)).toEqual(
      expect.arrayContaining([
        'rich-text',
        'hidden-column',
        'frozen-rows',
        'formula:ACCRINT',
        'export-control-legend',
      ]),
    );
    expect(converted.cellProvenance!.B2.actor.id).toBe('system');
  });
  it('reads authored decimal XML directly instead of accepting rounded JavaScript numbers', async () => {
    const document = new ExcelJS.Workbook();
    document.addWorksheet('Sheet1').getCell('A1').value = 0;
    const files = unzipSync(new Uint8Array(await document.xlsx.writeBuffer()));
    files['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(files['xl/worksheets/sheet1.xml']).replace(
        '<v>0</v>',
        '<v>0.9999999999999999999999999</v>',
      ),
    );
    const imported = await convertXlsx(zipSync(files));
    expect(imported.workbook.worksheets[0].data[0][0]).toBe('0.9999999999999999999999999');
  });
  it('uses manual calculation and preserves existing formula results during export', async () => {
    const workbook = createWorkbook([['=1/3*3', '1E-25']]);
    const result = calculateWorkbook(workbook),
      bytes = await exportXlsx(workbook, result),
      files = unzipSync(bytes);
    expect(strFromU8(files['xl/workbook.xml'])).toContain('calcMode="manual"');
    expect(strFromU8(files['xl/workbook.xml'])).toContain('calcOnSave="0"');
    const opened = new ExcelJS.Workbook();
    await opened.xlsx.load(bytes as never);
    expect(opened.worksheets[0].getCell('A1').formula).toBe('1/3*3');
    expect(opened.worksheets[0].getCell('A1').result).toBe(1);
    expect(opened.worksheets[0].getCell('B1').value).toBe(1e-25);
  });
});

describe('Host SE-XL-3: bounded OOXML input', () => {
  it('rejects malformed inputs, XML entities and files over the configured input limit', async () => {
    expect(() => inspectXlsx(new Uint8Array([1, 2, 3]))).toThrow();
    const document = new ExcelJS.Workbook();
    document.addWorksheet('Sheet1').getCell('A1').value = 1;
    const bytes = new Uint8Array(await document.xlsx.writeBuffer());
    expect(() => inspectXlsx(bytes, { maxBytes: bytes.length - 1 })).toThrow();
    const files = unzipSync(bytes);
    files['xl/workbook.xml'] = strToU8(
      '<!DOCTYPE workbook [<!ENTITY unsafe SYSTEM "file:///etc/passwd">]>' +
        strFromU8(files['xl/workbook.xml']),
    );
    expect(() => inspectXlsx(zipSync(files))).toThrow();
    delete files['xl/workbook.xml'];
    files['../bad.xml'] = strToU8('bad');
    expect(() => inspectXlsx(zipSync(files))).toThrow();
  });
});

it('Host SE-XL-5.2: exports immutable operands, aliases and exact caches without live resolution', async () => {
  const workbook = createWorkbook([
    [
      '=WORKITEM.FIELD("mass")+1',
      '=(1 m)/(2 m)',
      '=EXPM1(1E-25)',
      '=STDEVS(1,2)',
      '=ACCRINT(1,2,3)',
      '=1-1E-25',
      '="EXPM1("',
    ],
  ]);
  let reads = 0;
  const result = calculateWorkbook(
    workbook,
    {},
    {
      callFunction: (name) => {
        reads++;
        return name === 'WORKITEM.FIELD'
          ? { kind: 'number', numerator: '2', denominator: '1', approximate: false }
          : undefined;
      },
    },
  );
  const before = reads;
  const bytes = await exportXlsx(workbook, result);
  expect(reads).toBe(before);
  const document = new ExcelJS.Workbook();
  await document.xlsx.load(bytes as never);
  const sheet = document.worksheets[0];
  expect(sheet.getCell('A1').formula).toBe('2+1');
  expect(sheet.getCell('A1').result).toBe(3);
  expect(sheet.getCell('B1').value).toBe(0.5);
  expect(sheet.getCell('C1').formula).toBeUndefined();
  expect(sheet.getCell('D1').formula).toBe('STDEV.S(1,2)');
  expect(sheet.getCell('E1').formula).toBe('ACCRINT(1,2,3)');
  expect(sheet.getCell('E1').result).toBeUndefined();
  expect(sheet.getCell('F1').result).toBe('0.9999999999999999999999999');
  expect(sheet.getCell('G1').formula).toBe('"EXPM1("');
});

it('Host SE-XL-1.2: resolves theme tints and indexed colors to authored hex styles', async () => {
  const document = new ExcelJS.Workbook(),
    sheet = document.addWorksheet('Colors');
  sheet.getCell('A1').value = 1;
  sheet.getCell('A1').font = { color: { theme: 4, tint: 0.5 } as Partial<ExcelJS.Color> };
  sheet.getCell('A2').value = 2;
  sheet.getCell('A2').font = { color: { indexed: 10 } as Partial<ExcelJS.Color> };
  const converted = await convertXlsx(new Uint8Array(await document.xlsx.writeBuffer()));
  expect(converted.workbook.worksheets[0].style!.A1).toMatch(/color:#[0-9a-f]{6}/);
  expect(converted.workbook.worksheets[0].style!.A2).toContain('color:#ff0000');
});
