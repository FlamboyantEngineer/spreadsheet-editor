import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { calculateWorkbook } from '../calculation/evaluator.js';
import { formatCells } from '../model/editing.js';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import { validateFormulaCapabilities } from './capabilities.js';
import { capturedText } from './captured-text.js';
import type { ExcelFormulaCapabilities } from './contracts.js';
import { convertXlsx } from './convert.js';
import { inspectXlsx } from './guard.js';
import { decodeImportRequest, encodeImportRequest } from './protocol.js';
import { exportNumberFormat, importNumberFormat } from './styles.js';
import { exportXlsxWithReport } from './write.js';

const exportXlsx = async (...args: Parameters<typeof exportXlsxWithReport>) =>
  (await exportXlsxWithReport(...args)).bytes;

describe('SE-XLSX-1.1 and SE-XLSX-2.1: Excel conversion and loss reporting', () => {
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
    expect(requireDefined(converted.formats).B2).toMatchObject({ kind: 'percent', decimals: 2 });
    expect(requireDefined(converted.style).B2).toContain('font-weight: bold');
    expect(requireDefined(converted.style).B2).toContain('#175cd3');
    expect(requireDefined(converted.columns)[0].width).toBe(145);
    expect(requireDefined(converted.rows)['1'].height).toBe(32);
    expect(requireDefined(converted.mergeCells).A4).toEqual([2, 1]);
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
    expect(requireDefined(converted.cellProvenance).B2.actor.id).toBe('system');
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

describe('SE-XLSX-1.1: bounded OOXML input', () => {
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

it('SE-XLSX-2.1: exports immutable operands, aliases and exact caches without live resolution', async () => {
  const workbook = createWorkbook([
    [
      '=CAPTURED.VALUE("mass")+1',
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
        return name === 'CAPTURED.VALUE'
          ? { kind: 'number', numerator: '2', denominator: '1', approximate: false }
          : undefined;
      },
    },
  );
  const before = reads;
  const bytes = await exportXlsx(workbook, result, {
    formulaCapabilities: { hostFunctions: [{ name: 'CAPTURED.VALUE' }] },
  });
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

it('SE-XLSX-1.1: resolves theme tints and indexed colors to authored hex styles', async () => {
  const document = new ExcelJS.Workbook(),
    sheet = document.addWorksheet('Colors');
  sheet.getCell('A1').value = 1;
  sheet.getCell('A1').font = { color: { theme: 4, tint: 0.5 } as Partial<ExcelJS.Color> };
  sheet.getCell('A2').value = 2;
  sheet.getCell('A2').font = { color: { indexed: 10 } as Partial<ExcelJS.Color> };
  const converted = await convertXlsx(new Uint8Array(await document.xlsx.writeBuffer()));
  expect(requireDefined(converted.workbook.worksheets[0].style).A1).toMatch(/color: #[0-9a-f]{6}/);
  expect(requireDefined(converted.workbook.worksheets[0].style).A2).toContain('color: #ff0000');
});

const exampleCapabilities: ExcelFormulaCapabilities = {
  hostFunctions: [
    'CAPTURED',
    'CAPTURED.VALUE',
    'CAPTURED.VALUE.LABEL',
    'CAPTURED.VALUE.PROMPT',
    'RECORD.FIELD',
    'PROFILE.NAME',
    'PROFILE.EMAIL',
    'PROFILE.FIRSTNAME',
    'PROFILE.LASTNAME',
  ].map((name) => ({
    name,
    capturedText: name.startsWith('CAPTURED') ? 'html-to-text' : 'plain',
  })),
};

describe('SE-XLSX-3.1: explicit serializable host capabilities', () => {
  it('requires explicit host descriptors during import, including aliases and every example function', async () => {
    const document = new ExcelJS.Workbook(),
      sheet = document.addWorksheet('Inputs');
    exampleCapabilities.hostFunctions.forEach(({ name }, index) => {
      sheet.getCell(1, index + 1).value = { formula: `${name}("x")`, result: 999 };
    });
    sheet.getCell(2, 1).value = { formula: 'ALIAS("x")', result: 99 };
    const bytes = new Uint8Array(await document.xlsx.writeBuffer());
    const undeclared = await convertXlsx(bytes);
    expect(
      undeclared.report.warnings.filter((warning) => warning.code.startsWith('formula:')),
    ).toHaveLength(10);
    const declared = await convertXlsx(
      bytes,
      {},
      {
        hostFunctions: [...exampleCapabilities.hostFunctions, { name: 'FOO', aliases: ['ALIAS'] }],
      },
    );
    expect(
      declared.report.warnings.filter((warning) => warning.code.startsWith('formula:')),
    ).toHaveLength(0);
    expect(declared.workbook.worksheets[0].data[0][0]).toBe('=CAPTURED("x")');
  });

  it('inlines repeated generic calls by exact parser span, including dynamic arguments and names', async () => {
    const workbook = createWorkbook([['=FOO(A2)+BAR("other")+constant', '=FOO(FOO(A2))'], ['4']]);
    let reads = 0;
    const result = calculateWorkbook(
      workbook,
      {},
      {
        callFunction: (name, args) => {
          reads++;
          return name === 'FOO'
            ? args[0]
            : { kind: 'number', numerator: '7', denominator: '1', approximate: false };
        },
        readName: () => ({ kind: 'number', numerator: '3', denominator: '1', approximate: false }),
      },
    );
    const before = reads;
    const output = await exportXlsxWithReport(workbook, result, {
      formulaCapabilities: { hostFunctions: [{ name: 'foo', aliases: ['bar'] }] },
    });
    expect(reads).toBe(before);
    expect(output.report.warnings).toEqual([]);
    const document = new ExcelJS.Workbook();
    await document.xlsx.load(output.bytes as never);
    expect(document.worksheets[0].getCell('A1').formula).toBe('4+7+3');
    expect(document.worksheets[0].getCell('B1').formula).toBe('4');
  });

  it('keeps plain captured text by default and parses rich text only when explicitly declared', async () => {
    const workbook = createWorkbook([['=FOO()']]);
    const html =
      '<b>A &amp; &quot;B&quot;</b><br>line<script>hidden()</script><style>hide</style><svg><text>omit</text></svg>';
    const result = calculateWorkbook(workbook, {}, { callFunction: () => html });
    const plain = await exportXlsx(workbook, result, {
      formulaCapabilities: { hostFunctions: [{ name: 'FOO' }] },
    });
    const rich = await exportXlsx(workbook, result, {
      formulaCapabilities: { hostFunctions: [{ name: 'FOO', capturedText: 'html-to-text' }] },
    });
    const p = new ExcelJS.Workbook(),
      r = new ExcelJS.Workbook();
    await p.xlsx.load(plain as never);
    await r.xlsx.load(rich as never);
    expect(p.worksheets[0].getCell('A1').formula).toBe('"' + html.replaceAll('"', '""') + '"');
    expect(r.worksheets[0].getCell('A1').formula).toBe('"A & ""B""\nline"');
    expect(r.worksheets[0].getCell('A1').result).toBe('A & "B"\nline');
  });

  it('projects long direct rich-call results as plain stored text when a formula literal cannot fit', async () => {
    const workbook = createWorkbook([['=ALIAS()']]);
    const expected = 'readable '.repeat(40) + '\nnext';
    const result = calculateWorkbook(
      workbook,
      {},
      {
        callFunction: () => `<b>${'readable '.repeat(40)}</b><br>next<script>omitted</script>`,
      },
    );
    const output = await exportXlsxWithReport(workbook, result, {
      formulaCapabilities: {
        hostFunctions: [{ name: 'FOO', aliases: ['ALIAS'], capturedText: 'html-to-text' }],
      },
    });
    const document = new ExcelJS.Workbook();
    await document.xlsx.load(output.bytes as never);
    expect(document.worksheets[0].getCell('A1').value).toBe(expected);
    expect(output.report.warnings).toContainEqual(
      expect.objectContaining({ code: 'captured-text', disposition: 'stored-value' }),
    );
  });

  it('reports unsupported, missing, mismatched and overlong captured values without live lookup', async () => {
    const workbook = createWorkbook([['=FOO()', '=FOO()', '=FOO()']]);
    const result = calculateWorkbook(
      workbook,
      {},
      { callFunction: (_name, _args, context) => (context.x === 2 ? 'x'.repeat(256) : 'stored') },
    );
    const undeclared = await exportXlsxWithReport(workbook, result);
    expect(
      undeclared.report.warnings.find((warning) => warning.code === 'unsupported-formula')?.count,
    ).toBe(3);
    const operands = Object.values(requireDefined(result.context).operands);
    delete requireDefined(result.context).operands[
      Object.keys(requireDefined(result.context).operands)[0]
    ];
    operands[1].formula = '=different';
    const captured = await exportXlsxWithReport(workbook, result, {
      formulaCapabilities: { hostFunctions: [{ name: 'FOO' }] },
    });
    expect(
      captured.report.warnings.find((warning) => warning.code === 'captured-operand')?.count,
    ).toBe(2);
    expect(
      captured.report.warnings.find((warning) => warning.code === 'captured-text')?.count,
    ).toBe(1);
    const document = new ExcelJS.Workbook();
    await document.xlsx.load(captured.bytes as never);
    expect(document.worksheets[0].getCell('A1').value).toBe('stored');
    expect(document.worksheets[0].getCell('C1').value).toBe('x'.repeat(256));
  });

  it('rejects malformed descriptors at parent and child request boundaries', () => {
    for (const value of [
      null,
      [],
      { hostFunctions: Array(129).fill({ name: 'FOO' }) },
      { hostFunctions: [{ name: 'BAD()' }] },
      { hostFunctions: [{ name: 'FOO', capturedText: 'execute' }] },
      { hostFunctions: [{ name: 'FOO', aliases: Array(17).fill('bar') }] },
      { hostFunctions: [{ name: 'FOO' }, { name: 'foo' }] },
      { hostFunctions: [{ name: 'FOO', aliases: ['SUM'] }] },
      { hostFunctions: [{ name: 'FOO', resolver: () => 1 }] },
    ])
      expect(() => validateFormulaCapabilities(value)).toThrow();
    const encoded = encodeImportRequest(new Uint8Array([1, 2, 3]), exampleCapabilities);
    const decoded = decodeImportRequest(encoded);
    expect(decoded.bytes).toEqual(Buffer.from([1, 2, 3]));
    expect(decoded.formulaCapabilities).toEqual(validateFormulaCapabilities(exampleCapabilities));
    encoded.writeUInt32LE(65_537);
    expect(() => decodeImportRequest(encoded)).toThrow('header length');
    expect(() => encodeImportRequest(new Uint8Array(20 * 1024 * 1024 + 1))).toThrow('byte limit');
  });

  it('bounds established HTML parsing and rejects unrepresentable strings', () => {
    expect(capturedText('<p>&copy; &#x1F600;<BR/>next</p><!-- ignored -->', 'html-to-text')).toBe(
      '© 😀\nnext',
    );
    expect(capturedText('<div title="x > y">safe</div>', 'html-to-text')).toBe('safe');
    expect(() => capturedText('x'.repeat(65_537), 'html-to-text')).toThrow('input limit');
    expect(() =>
      capturedText('<i>'.repeat(130) + 'x' + '</i>'.repeat(130), 'html-to-text'),
    ).toThrow('tree limit');
    expect(() => capturedText('bad\u0001', 'plain')).toThrow('cannot represent');
  });
});

describe('SE-XLSX-4.1: locale, exact decimals and explicit metadata losses', () => {
  it('round trips EUR/de-DE and USD/en-US format metadata without rounding stored decimals', async () => {
    const formats = [
      { kind: 'currency' as const, currency: 'EUR', locale: 'de-DE', decimals: 2 },
      { kind: 'currency' as const, currency: 'USD', locale: 'en-US', decimals: 2 },
      { kind: 'percent' as const, locale: 'de-DE', decimals: 3 },
      { kind: 'date' as const, locale: 'de-DE' },
    ];
    for (const format of formats)
      expect(importNumberFormat(exportNumberFormat(format)).format).toMatchObject(format);
    const workbook = createWorkbook([['9007199254740993', '-1234.567', '0.125', '60']]);
    workbook.worksheets[0].formats = Object.fromEntries(
      formats.map((format, index) => [String.fromCharCode(65 + index) + '1', format]),
    );
    const output = await exportXlsxWithReport(workbook, calculateWorkbook(workbook));
    expect(output.report.warnings).toEqual([]);
    const imported = await convertXlsx(output.bytes);
    expect(imported.workbook.worksheets[0].data[0]).toEqual([
      "'9007199254740993",
      '-1234.567',
      '0.125',
      '60',
    ]);
    formats.forEach((format, index) =>
      expect(
        imported.workbook.worksheets[0].formats?.[String.fromCharCode(65 + index) + '1'],
      ).toMatchObject(format),
    );
  });
  it('reports a locale that has no representable Excel LCID rather than inventing one', async () => {
    const workbook = createWorkbook([['1']]);
    workbook.worksheets[0].formats = {
      A1: { kind: 'number', locale: 'en-001' },
    } as (typeof workbook.worksheets)[0]['formats'];
    const output = await exportXlsxWithReport(workbook, calculateWorkbook(workbook));
    expect(output.report.warnings).toContainEqual(
      expect.objectContaining({ code: 'number-format', disposition: 'not-exported' }),
    );
  });
});

it('SE-XLSX-4.1: keeps exact numeric text beyond display precision instead of rounding it', async () => {
  const workbook = createWorkbook([['1']]),
    result = calculateWorkbook(workbook);
  const numerator = '1' + '0'.repeat(10000) + '1';
  result.worksheets[0].values[0][0] = {
    kind: 'number',
    numerator,
    denominator: '1',
    approximate: false,
  };
  const output = await exportXlsx(workbook, result);
  const opened = new ExcelJS.Workbook();
  await opened.xlsx.load(output as never);
  expect(opened.worksheets[0].getCell('A1').value).toBe(numerator);
});

it('SE-XLSX-4.1: projects canonical quantities once, preserves temperature differences and reports typed losses', async () => {
  const workbook = createWorkbook([
    ['alternate authored text', 'another representation', '2'],
    ['=30 °C-25 °C'],
  ]);
  const sheet = workbook.worksheets[0];
  sheet.inputValues = {
    A1: {
      kind: 'number',
      numerator: '5',
      denominator: '1',
      approximate: false,
      dimension: [0, 0, 0, 0, 1, 0, 0],
      temperature: 'difference',
      provenance: { 'constant-set': 'rev1' },
    },
    B1: {
      kind: 'number',
      numerator: '27315',
      denominator: '100',
      approximate: true,
      precision: 80,
      dimension: [0, 0, 0, 0, 1, 0, 0],
      temperature: 'absolute',
    },
    C1: {
      kind: 'number',
      numerator: '1',
      denominator: '1',
      approximate: false,
      quantityKind: 'angle',
    },
  };
  sheet.formats = {
    A1: { kind: 'unit', unit: '°F', inputUnit: '°C' },
    B1: { kind: 'unit', unit: '°C', inputUnit: '°F' },
    A2: { kind: 'unit', unit: '°C' },
  };
  const result = calculateWorkbook(workbook);
  const output = await exportXlsxWithReport(workbook, result);
  const opened = new ExcelJS.Workbook();
  await opened.xlsx.load(output.bytes as never);
  expect(opened.worksheets[0].getCell('A1').value).toBe('9 °F');
  expect(opened.worksheets[0].getCell('B1').value).toBe('≈0 °C');
  expect(opened.worksheets[0].getCell('C1').value).toBe('1 rad');
  expect(opened.worksheets[0].getCell('A2').value).toBe('5 °C');
  expect(output.report.warnings.find((warning) => warning.code === 'canonical-input')?.count).toBe(
    3,
  );
  expect(
    output.report.warnings.find((warning) => warning.code === 'quantity-metadata')?.count,
  ).toBe(4);
  expect(output.report.formulas).toBe(1);
  const imported = await convertXlsx(output.bytes);
  expect(calculateWorkbook(imported.workbook).worksheets[0].values[0]).toEqual([
    '9 °F',
    '≈0 °C',
    '1 rad',
  ]);
});

it('SE-XLSX-4.1: rejects malformed quantity and provenance result metadata at the export boundary', async () => {
  const workbook = createWorkbook([['1']]);
  for (const metadata of [
    { temperature: 'relative' },
    { provenance: { ns: 3 } },
    { dimension: [1, 2] },
  ]) {
    const result = calculateWorkbook(workbook);
    Object.assign(requireDefined(result.worksheets[0].values[0][0]), metadata);
    await expect(exportXlsx(workbook, result)).rejects.toThrow();
  }
});

it('SE-CHROME-1.1: an imported style and the same style applied in the editor store the same text', async () => {
  const document = new ExcelJS.Workbook(),
    sheet = document.addWorksheet('Styles');
  sheet.getCell('A1').value = 'Label';
  sheet.getCell('A1').font = { bold: true };
  const imported = (await convertXlsx(new Uint8Array(await document.xlsx.writeBuffer()))).workbook;
  const style = requireDefined(imported.worksheets[0].style).A1;
  expect(style).toBe('font-weight: bold; text-align: left');
  const cell = { x1: 0, y1: 0, x2: 0, y2: 0 };
  const plain = { ...imported, worksheets: [{ ...imported.worksheets[0], style: {} }] };
  const edited = formatCells(
    plain,
    { sheetId: imported.worksheets[0].id, range: cell },
    {
      properties: {
        'font-weight': 'bold',
        'text-align': 'left',
      },
    },
  );
  expect(requireDefined(edited.worksheets[0].style).A1).toBe(style);
  // Re-applying the imported properties leaves the stored text unchanged.
  expect(
    requireDefined(
      formatCells(
        imported,
        { sheetId: imported.worksheets[0].id, range: cell },
        { properties: { 'font-weight': 'bold' } },
      ).worksheets[0].style,
    ).A1,
  ).toBe(style);
});
