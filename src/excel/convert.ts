import ExcelJS from 'exceljs';
import { createArithmetic } from '../calculation/rational.js';
import { requireDefined } from '../model/require-defined.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetRange,
} from '../model/spreadsheet-addressing.js';
import { assertWorkbook, type Workbook, type Worksheet } from '../model/workbook.js';
import { hostFunctionMap, validateFormulaCapabilities } from './capabilities.js';
import { readThemeColors } from './colors.js';
import type { ExcelFormulaCapabilities, XlsxImport, XlsxReport } from './contracts.js';
import { inspectXlsx, type XlsxLimits } from './guard.js';
import { importCell, type ImportContext, type SheetImport } from './import-cell.js';
import { createWarnings } from './warnings.js';

const decode = new TextDecoder();
function attrs(tag: string): Record<string, string> {
  return Object.fromEntries(
    [...tag.matchAll(/([\w:]+)="([^"]*)"/g)].map((match) => [match[1], match[2]]),
  );
}
type Parts = Record<string, Uint8Array>;
const text = (parts: Parts, name: string) => decode.decode(parts[name] ?? new Uint8Array());

/** The worksheet part paths, in workbook order, from the workbook and its relationships. */
function worksheetPaths(parts: Parts): (string | undefined)[] {
  const relationships = new Map(
    [...text(parts, 'xl/_rels/workbook.xml.rels').matchAll(/<Relationship\b[^>]*>/g)].map(
      (match) => {
        const a = attrs(match[0]);
        return [a.Id, a.Target];
      },
    ),
  );
  return [...decode.decode(parts['xl/workbook.xml']).matchAll(/<sheet\b[^>]*>/g)].map((match) =>
    relationships.get(attrs(match[0])['r:id']),
  );
}

/** The XML of a worksheet part; only `xl/worksheets/sheetN.xml` parts are read. */
function worksheetXml(parts: Parts, target: string | undefined): string {
  const part = target?.startsWith('/xl/')
    ? target.slice(1)
    : target?.startsWith('xl/')
      ? target
      : `xl/${target}`;
  if (!/^xl\/worksheets\/sheet[1-9]\d*\.xml$/.test(part) || !parts[part])
    throw new TypeError('Unsupported workbook part layout.');
  return decode.decode(parts[part]);
}

/** Plain numeric cells as the XML wrote them, so no binary float rounding reaches the model. */
function authoredNumbers(xml: string): Map<string, string> {
  const numbers = new Map<string, string>();
  for (const cell of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const a = attrs(cell[1]);
    if ((a.t && a.t !== 'n') || /<f(?:\s|>)/.test(cell[2])) continue;
    const number = /<v>([^<]+)<\/v>/.exec(cell[2])?.[1];
    if (number && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(number))
      numbers.set(a.r, number);
  }
  return numbers;
}

const emptySheet = (index: number, name: string): Worksheet => ({
  id: `sheet-${index + 1}`,
  name,
  data: [],
  columns: [],
  rows: {},
  style: {},
  formats: {},
  mergeCells: {},
  cellProvenance: {},
});

function importColumns(context: ImportContext, source: ExcelJS.Worksheet, sheet: Worksheet) {
  (source.columns ?? []).forEach((column, x) => {
    requireDefined(sheet.columns)[x] = { width: column.width ? column.width * 7 + 5 : 100 };
    if (column.hidden)
      context.warn({
        code: 'hidden-column',
        message: 'Hidden columns are imported visible.',
        location: `${source.name}!${x + 1}`,
        disposition: 'skipped',
      });
  });
}

function importRows(
  context: ImportContext,
  source: ExcelJS.Worksheet,
  { target, maxCells }: { target: SheetImport; maxCells: number },
) {
  const { sheet } = target;
  source.eachRow({ includeEmpty: false }, (row, y) => {
    if (y > maxCells)
      throw new RangeError('Worksheet extent exceeds the standalone editor row limit.');
    while (sheet.data.length < y) sheet.data.push([]);
    if (row.height) requireDefined(sheet.rows)[String(y - 1)] = { height: (row.height * 4) / 3 };
    if (row.hidden)
      context.warn({
        code: 'hidden-row',
        message: 'Hidden rows are imported visible.',
        location: `${source.name}!${y}`,
        disposition: 'skipped',
      });
    row.eachCell({ includeEmpty: false }, (cell, x) => importCell(context, target, { cell, x, y }));
  });
}

function importLayout(context: ImportContext, source: ExcelJS.Worksheet, sheet: Worksheet) {
  for (const range of source.model.merges ?? []) {
    const parsed = parseSpreadsheetRange(range);
    if (!parsed) continue;
    requireDefined(sheet.mergeCells)[formatSpreadsheetAddress(parsed.start.x, parsed.start.y)] = [
      parsed.end.x - parsed.start.x + 1,
      parsed.end.y - parsed.start.y + 1,
    ];
    context.report.mergedRanges++;
  }
  for (const view of source.views ?? []) {
    if (view.state !== 'frozen') continue;
    if (view.xSplit) sheet.freezeColumns = view.xSplit;
    if (view.ySplit)
      context.warn({
        code: 'frozen-rows',
        message: 'Frozen rows are not imported.',
        location: source.name,
        disposition: 'skipped',
      });
  }
}

const SHEET_FEATURES = [
  [
    /<conditionalFormatting\b/,
    'conditional-formatting',
    'Conditional formatting is not imported.',
    'not-imported',
  ],
  [/<dataValidation\b/, 'validation', 'Data validation is not imported.', 'not-imported'],
  [/<pageSetup\b/, 'page-setup', 'Page setup is not imported.', 'skipped'],
] as const;

function featureWarnings(context: ImportContext, source: ExcelJS.Worksheet, xml: string) {
  const location = source.name;
  if (source.headerFooter && Object.values(source.headerFooter).some(Boolean))
    context.warn({
      code: 'print-state',
      message: 'Print headers, footers and page setup are not imported.',
      location,
      disposition: 'skipped',
    });
  for (const [pattern, code, message, disposition] of SHEET_FEATURES)
    if (pattern.test(xml)) context.warn({ code, message, location, disposition });
}

function importWorksheet(
  context: ImportContext,
  source: ExcelJS.Worksheet,
  { index, xml, maxCells }: { index: number; xml: string; maxCells: number },
): Worksheet {
  const sheet = emptySheet(index, source.name);
  const target = { sheet, name: source.name, authoredNumbers: authoredNumbers(xml) };
  if (source.state !== 'visible')
    context.warn({
      code: 'hidden-worksheet',
      message: 'Hidden worksheets are imported visible.',
      location: source.name,
      disposition: 'skipped',
    });
  importColumns(context, source, sheet);
  importRows(context, source, { target, maxCells });
  if (!sheet.data.length) sheet.data = [['']];
  importLayout(context, source, sheet);
  featureWarnings(context, source, xml);
  return sheet;
}

const PART_CATEGORIES: readonly [RegExp, string][] = [
  [/^xl\/charts\/chart\d+\.xml$/, 'charts'],
  [/^xl\/media\//, 'images'],
  [/^xl\/pivotTables\//, 'pivot-tables'],
  [/^xl\/tables\/table\d+\.xml$/, 'tables'],
  [/^xl\/externalLinks\/externalLink\d+\.xml$/, 'external-links'],
  [/^xl\/threadedComments\//, 'threaded-comments'],
];

function partWarnings(context: ImportContext, parts: Parts) {
  for (const name of Object.keys(parts)) {
    const category = PART_CATEGORIES.find(([pattern]) => pattern.test(name))?.[1];
    if (category)
      context.warn({
        code: category,
        message: `${category.replaceAll('-', ' ')} are not imported; cell values remain.`,
        location: name,
      });
  }
}

/** Runs only in the isolated import child. Excel's formula caches are never read. */
export async function convertXlsx(
  bytes: Uint8Array,
  limits: XlsxLimits = {},
  formulaCapabilities?: ExcelFormulaCapabilities,
): Promise<XlsxImport> {
  const external = hostFunctionMap(validateFormulaCapabilities(formulaCapabilities));
  const { parts } = inspectXlsx(bytes, limits),
    warnings = createWarnings('not-imported');
  const theme = readThemeColors(text(parts, 'xl/theme/theme1.xml'));
  const document = new ExcelJS.Workbook();
  await document.xlsx.load(bytes as never);
  const workbook: Workbook = { schemaVersion: 1, worksheets: [] };
  const report: XlsxReport = {
    worksheets: 0,
    cells: 0,
    formulas: 0,
    mergedRanges: 0,
    warnings: [],
  };
  const context: ImportContext = {
    warn: warnings.warn,
    report,
    theme,
    arithmetic: createArithmetic(),
    date1904: Boolean(document.properties.date1904),
    external,
  };
  const paths = worksheetPaths(parts),
    maxCells = limits.maxCells ?? 250000;
  document.worksheets.forEach((source, index) => {
    const xml = worksheetXml(parts, paths[index]);
    workbook.worksheets.push(importWorksheet(context, source, { index, xml, maxCells }));
    report.worksheets++;
  });
  partWarnings(context, parts);
  assertWorkbook(workbook, { maxCells, maxWorksheets: limits.maxWorksheets ?? 200 });
  report.warnings = warnings.list();
  return { workbook, report };
}
