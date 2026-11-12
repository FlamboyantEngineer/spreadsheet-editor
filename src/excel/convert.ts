import ExcelJS from 'exceljs';
import { readThemeColors } from './colors.js';
import { inspectXlsx, type XlsxLimits } from './guard.js';
import type { ExcelFormulaCapabilities, XlsxImport, XlsxWarning } from './contracts.js';
import { hostFunctionMap, validateFormulaCapabilities } from './capabilities.js';
import { importCellStyle, importNumberFormat } from './styles.js';
import { createArithmetic } from '../calculation/rational.js';
import { formatCellValue } from '../calculation/format.js';
import { SUPPORTED_FUNCTIONS } from '../calculation/evaluator.js';
import { parseFormula, type FormulaNode } from '../model/formula-parser.js';
import {
  assertWorkbook,
  type CellInput,
  type Workbook,
  type Worksheet,
} from '../model/workbook.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetRange,
} from '../model/spreadsheet-addressing.js';

const decode = new TextDecoder();
function unsupported(node: FormulaNode, external: ReadonlyMap<string, unknown>): string[] {
  if (node.kind === 'call')
    return [
      ...(!SUPPORTED_FUNCTIONS.includes(node.name) && !external.has(node.name) ? [node.name] : []),
      ...node.args.flatMap((arg) => unsupported(arg, external)),
    ];
  if (node.kind === 'binary')
    return [...unsupported(node.left, external), ...unsupported(node.right, external)];
  return node.kind === 'unary' ? unsupported(node.operand, external) : [];
}
function attrs(tag: string): Record<string, string> {
  return Object.fromEntries(
    [...tag.matchAll(/([\w:]+)="([^"]*)"/g)].map((match) => [match[1], match[2]]),
  );
}

/** Runs only in the isolated import child. Excel's formula caches are never read. */
export async function convertXlsx(
  bytes: Uint8Array,
  limits: XlsxLimits = {},
  formulaCapabilities?: ExcelFormulaCapabilities,
): Promise<XlsxImport> {
  const external = hostFunctionMap(validateFormulaCapabilities(formulaCapabilities));
  const inspected = inspectXlsx(bytes, limits),
    warnings = new Map<string, XlsxWarning>();
  const warn = (
    code: string,
    message: string,
    location: string,
    disposition: XlsxWarning['disposition'] = 'not-imported',
  ) => {
    const entry = warnings.get(code) ?? { code, message, disposition, count: 0, locations: [] };
    entry.count++;
    if (entry.locations.length < 5 && !entry.locations.includes(location))
      entry.locations.push(location);
    warnings.set(code, entry);
  };
  const theme = readThemeColors(
    decode.decode(inspected.parts['xl/theme/theme1.xml'] ?? new Uint8Array()),
  );
  const document = new ExcelJS.Workbook();
  await document.xlsx.load(bytes as never);
  const workbook: Workbook = { schemaVersion: 1, worksheets: [] };
  const report = {
    worksheets: 0,
    cells: 0,
    formulas: 0,
    mergedRanges: 0,
    warnings: [] as XlsxWarning[],
  };
  const relationships = new Map(
    [
      ...decode
        .decode(inspected.parts['xl/_rels/workbook.xml.rels'] ?? new Uint8Array())
        .matchAll(/<Relationship\b[^>]*>/g),
    ].map((match) => {
      const a = attrs(match[0]);
      return [a.Id, a.Target];
    }),
  );
  const paths = [
    ...decode.decode(inspected.parts['xl/workbook.xml']).matchAll(/<sheet\b[^>]*>/g),
  ].map((match) => relationships.get(attrs(match[0])['r:id']));
  const arithmetic = createArithmetic();
  document.worksheets.forEach((source, index) => {
    const target = paths[index];
    const part = target?.startsWith('/xl/')
      ? target.slice(1)
      : target?.startsWith('xl/')
        ? target
        : `xl/${target}`;
    if (!part || !/^xl\/worksheets\/sheet[1-9]\d*\.xml$/.test(part) || !inspected.parts[part])
      throw new TypeError('Unsupported workbook part layout.');
    const xml = decode.decode(inspected.parts[part]);
    const authoredNumbers = new Map<string, string>();
    for (const cell of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const a = attrs(cell[1]);
      if ((a.t && a.t !== 'n') || /<f(?:\s|>)/.test(cell[2])) continue;
      const number = /<v>([^<]+)<\/v>/.exec(cell[2])?.[1];
      if (number && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(number))
        authoredNumbers.set(a.r, number);
    }
    const sheet: Worksheet = {
      id: `sheet-${index + 1}`,
      name: source.name,
      data: [],
      columns: [],
      rows: {},
      style: {},
      formats: {},
      mergeCells: {},
      cellProvenance: {},
    };
    if (source.state !== 'visible')
      warn('hidden-worksheet', 'Hidden worksheets are imported visible.', source.name, 'skipped');
    (source.columns ?? []).forEach((column, x) => {
      sheet.columns![x] = { width: column.width ? column.width * 7 + 5 : 100 };
      if (column.hidden)
        warn(
          'hidden-column',
          'Hidden columns are imported visible.',
          `${source.name}!${x + 1}`,
          'skipped',
        );
    });
    source.eachRow({ includeEmpty: false }, (row, y) => {
      if (y > (limits.maxCells ?? 250000))
        throw new RangeError('Worksheet extent exceeds the standalone editor row limit.');
      while (sheet.data.length < y) sheet.data.push([]);
      if (row.height) sheet.rows![String(y - 1)] = { height: (row.height * 4) / 3 };
      if (row.hidden)
        warn('hidden-row', 'Hidden rows are imported visible.', `${source.name}!${y}`, 'skipped');
      row.eachCell({ includeEmpty: false }, (cell, x) => {
        if (cell.isMerged && cell.master.address !== cell.address) return;
        const address = formatSpreadsheetAddress(x - 1, y - 1),
          location = `${source.name}!${address}`;
        let value: CellInput;
        const format = importNumberFormat(cell.numFmt);
        if (!format)
          warn(
            `number-format:${cell.numFmt}`,
            'Unsupported number format falls back to General.',
            location,
          );
        if (format && format.kind !== 'general') sheet.formats![address] = format;
        if (cell.formula) {
          value = '=' + cell.formula;
          report.formulas++;
          try {
            for (const name of unsupported(parseFormula(cell.formula), external))
              warn(
                `formula:${name}`,
                `${name} is retained but is unavailable in the calculation catalog.`,
                location,
              );
          } catch {
            warn(
              'malformed-formula',
              'Formula is retained as authored but cannot be parsed.',
              location,
            );
          }
        } else if (authoredNumbers.has(address)) {
          value = authoredNumbers.get(address)!;
          if (
            document.properties.date1904 &&
            format &&
            ['date', 'time', 'datetime'].includes(format.kind)
          )
            value = formatCellValue(
              arithmetic.add(arithmetic.parse(value), arithmetic.parse('1462')),
              { significantDigits: 10000, notation: 'decimal' },
            );
        } else if (cell.value instanceof Date) {
          // Numeric date serials normally use the raw XML path above.
          value = String((cell.value.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
        } else if (typeof cell.value === 'object' && cell.value !== null) {
          if ('richText' in cell.value) {
            value = cell.value.richText.map((part) => part.text).join('');
            warn('rich-text', 'Rich text becomes plain text.', location);
          } else if ('hyperlink' in cell.value) {
            value = cell.value.text;
            warn('hyperlink', 'Hyperlinks become plain text.', location);
          } else if ('error' in cell.value) value = cell.value.error;
          else value = cell.text;
        } else value = cell.value as CellInput;
        const values = sheet.data[y - 1];
        while (values.length < x) values.push(null);
        values[x - 1] = value;
        const style = importCellStyle(cell.style, theme);
        const alignment = cell.alignment?.horizontal
          ? ''
          : typeof value === 'boolean' || (typeof value === 'string' && value.startsWith('#'))
            ? 'text-align:center'
            : authoredNumbers.has(address)
              ? 'text-align:right'
              : 'text-align:left';
        sheet.style![address] = [style, alignment].filter(Boolean).join(';');
        sheet.cellProvenance![address] = {
          actor: { id: 'system', name: 'System' },
          source: 'system',
        };
        report.cells++;
        if (cell.note) warn('note', 'Excel notes and comments are not imported.', location);
        if (
          cell.fill?.type === 'pattern' &&
          cell.fill.pattern !== 'solid' &&
          cell.fill.pattern !== 'none'
        )
          warn('pattern-fill', 'Pattern fills are not imported.', location);
        if (cell.alignment?.textRotation)
          warn('rotated-text', 'Text rotation is not imported.', location);
        if (typeof value === 'string' && /\b(?:ITAR|EAR99|ECCN|EXPORT CONTROLLED)\b/i.test(value))
          warn(
            'export-control-legend',
            'Export-control text is retained; set the artifact marking in the host application.',
            location,
          );
      });
    });
    if (!sheet.data.length) sheet.data = [['']];
    for (const range of source.model.merges ?? []) {
      const parsed = parseSpreadsheetRange(range);
      if (parsed) {
        sheet.mergeCells![formatSpreadsheetAddress(parsed.start.x, parsed.start.y)] = [
          parsed.end.x - parsed.start.x + 1,
          parsed.end.y - parsed.start.y + 1,
        ];
        report.mergedRanges++;
      }
    }
    for (const view of source.views ?? [])
      if (view.state === 'frozen') {
        if (view.xSplit) sheet.freezeColumns = view.xSplit;
        if (view.ySplit)
          warn('frozen-rows', 'Frozen rows are not imported.', source.name, 'skipped');
      }
    if (source.headerFooter && Object.values(source.headerFooter).some(Boolean))
      warn(
        'print-state',
        'Print headers, footers and page setup are not imported.',
        source.name,
        'skipped',
      );
    for (const [pattern, code, message, disposition] of [
      [
        /<conditionalFormatting\b/g,
        'conditional-formatting',
        'Conditional formatting is not imported.',
        'not-imported',
      ],
      [/<dataValidation\b/g, 'validation', 'Data validation is not imported.', 'not-imported'],
      [/<pageSetup\b/g, 'page-setup', 'Page setup is not imported.', 'skipped'],
    ] as const)
      if (pattern.test(xml)) warn(code, message, source.name, disposition);
    workbook.worksheets.push(sheet);
    report.worksheets++;
  });
  for (const name of Object.keys(inspected.parts)) {
    const category = /^xl\/charts\/chart\d+\.xml$/.test(name)
      ? 'charts'
      : /^xl\/media\//.test(name)
        ? 'images'
        : /^xl\/pivotTables\//.test(name)
          ? 'pivot-tables'
          : /^xl\/tables\/table\d+\.xml$/.test(name)
            ? 'tables'
            : /^xl\/externalLinks\/externalLink\d+\.xml$/.test(name)
              ? 'external-links'
              : /^xl\/threadedComments\//.test(name)
                ? 'threaded-comments'
                : null;
    if (category)
      warn(
        category,
        `${category.replaceAll('-', ' ')} are not imported; cell values remain.`,
        name,
      );
  }
  assertWorkbook(workbook, {
    maxCells: limits.maxCells ?? 250000,
    maxWorksheets: limits.maxWorksheets ?? 200,
  });
  report.warnings = [...warnings.values()];
  return { workbook, report };
}
