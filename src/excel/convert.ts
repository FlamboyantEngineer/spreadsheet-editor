import ExcelJS from 'exceljs';
import { createArithmetic } from '../calculation/rational.js';
import { requireDefined } from '../model/require-defined.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetAddress,
  parseSpreadsheetRange,
} from '../model/spreadsheet-addressing.js';
import { assertWorkbook, type Workbook, type Worksheet } from '../model/workbook.js';
import { hostFunctionMap, validateFormulaCapabilities } from './capabilities.js';
import { readThemeColors } from './colors.js';
import type { ExcelFormulaCapabilities, XlsxImport, XlsxReport } from './contracts.js';
import { inspectXlsx, type XlsxLimits } from './guard.js';
import { importCell, type ImportContext } from './import-cell.js';
import { createWarnings } from './warnings.js';
import type { WorksheetXml } from './worksheet-xml.js';
import { readXlsxXml } from './xlsx-xml.js';

const decode = new TextDecoder();
type Parts = Record<string, Uint8Array>;
const text = (parts: Parts, name: string) => decode.decode(parts[name] ?? new Uint8Array());

/** The worksheet part paths, in workbook order, from the workbook and its relationships. */
function worksheetPaths(parts: Parts): (string | undefined)[] {
  const relationships = new Map<string, string>();
  const paths: (string | undefined)[] = [];
  readXlsxXml(parts['xl/_rels/workbook.xml.rels'], {
    open(name, attributes, depth) {
      if (depth === 2 && name === 'Relationship')
        relationships.set(attributes.Id, attributes.Target);
    },
  });
  readXlsxXml(parts['xl/workbook.xml'], {
    open(name, attributes, depth) {
      if (depth === 3 && name === 'sheet') paths.push(relationships.get(attributes['r:id']));
    },
  });
  return paths;
}

/** The XML of a worksheet part; only `xl/worksheets/sheetN.xml` parts are read. */
function worksheetXml(parts: ReadonlyMap<string, WorksheetXml>, target: string | undefined) {
  const part = target?.startsWith('/xl/')
    ? target.slice(1)
    : target?.startsWith('xl/')
      ? target
      : `xl/${target}`;
  const worksheet = parts.get(part);
  if (!worksheet) throw new TypeError('Unsupported workbook part layout.');
  return worksheet;
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

function importRowGeometry(
  context: ImportContext,
  source: ExcelJS.Worksheet,
  { sheet, rows }: { sheet: Worksheet; rows: readonly number[] },
) {
  for (const y of rows) {
    const row = source.getRow(y);
    if (row.height) requireDefined(sheet.rows)[String(y - 1)] = { height: (row.height * 4) / 3 };
    if (row.hidden)
      context.warn({
        code: 'hidden-row',
        message: 'Hidden rows are imported visible.',
        location: `${source.name}!${y}`,
        disposition: 'skipped',
      });
  }
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
    'conditionalFormatting',
    'conditional-formatting',
    'Conditional formatting is not imported.',
    'not-imported',
  ],
  ['dataValidations', 'validation', 'Data validation is not imported.', 'not-imported'],
  ['pageSetup', 'page-setup', 'Page setup is not imported.', 'skipped'],
] as const;

function featureWarnings(
  context: ImportContext,
  source: ExcelJS.Worksheet,
  features: ReadonlySet<string>,
) {
  const location = source.name;
  if (source.headerFooter && Object.values(source.headerFooter).some(Boolean))
    context.warn({
      code: 'print-state',
      message: 'Print headers, footers and page setup are not imported.',
      location,
      disposition: 'skipped',
    });
  for (const [element, code, message, disposition] of SHEET_FEATURES)
    if (features.has(element)) context.warn({ code, message, location, disposition });
}

function importWorksheet(
  context: ImportContext,
  source: ExcelJS.Worksheet,
  { index, xml, maxCells }: { index: number; xml: WorksheetXml; maxCells: number },
): Worksheet {
  const sheet = emptySheet(index, source.name);
  if (source.state !== 'visible')
    context.warn({
      code: 'hidden-worksheet',
      message: 'Hidden worksheets are imported visible.',
      location: source.name,
      disposition: 'skipped',
    });
  importColumns(context, source, sheet);
  importRowGeometry(context, source, { sheet, rows: xml.rows });
  for (const [address, number] of xml.cells) {
    const { x, y } = requireDefined(parseSpreadsheetAddress(address));
    if ((x + 1) * (y + 1) > maxCells)
      throw new RangeError('Worksheet extent exceeds the standalone editor cell limit.');
    importCell(context, sheet, { cell: source.getCell(address), address, x, y, number });
  }
  if (!sheet.data.length) sheet.data = [['']];
  importLayout(context, source, sheet);
  featureWarnings(context, source, xml.features);
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
  const { parts, worksheets } = inspectXlsx(bytes, limits),
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
    const xml = worksheetXml(worksheets, paths[index]);
    workbook.worksheets.push(importWorksheet(context, source, { index, xml, maxCells }));
    report.worksheets++;
  });
  partWarnings(context, parts);
  assertWorkbook(workbook, { maxCells, maxWorksheets: limits.maxWorksheets ?? 200 });
  report.warnings = warnings.list();
  return { workbook, report };
}
