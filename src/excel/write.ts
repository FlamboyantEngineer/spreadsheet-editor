import ExcelJS from 'exceljs';
import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { assertWorkbook, type Workbook } from '../model/workbook.js';
import { parseCellStyle } from '../model/editing.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetAddress,
} from '../model/spreadsheet-addressing.js';
import { parseFormula, type FormulaNode } from '../model/formula-parser.js';
import { createArithmetic } from '../calculation/rational.js';
import { formatCellValue } from '../calculation/format.js';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import { isNumeric, isError, type CellValue } from '../calculation/values.js';
import {
  SUPPORTED_FUNCTIONS,
  formulaOperandKey,
  type CalculationResult,
} from '../calculation/evaluator.js';
import { exportCellStyle, exportNumberFormat } from './styles.js';

const exact = createArithmetic({ maxIntegerDigits: 100000 });
function excelValue(value: CellValue): ExcelJS.CellValue {
  if (isError(value))
    return ['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'].includes(value.code)
      ? { error: value.code as ExcelJS.CellErrorValue['error'] }
      : value.code;
  if (!isNumeric(value)) return value;
  if (value.dimension) return formatCellValue(value, { significantDigits: 10000 });
  const rendered = formatCellValue(value, { significantDigits: 10000, approximationMarker: false }),
    native = Number(rendered);
  if (Number.isFinite(native) && exact.compare(exact.parse(String(native)), value) === 0)
    return native;
  let denominator = BigInt(value.denominator);
  while (denominator % 2n === 0n) denominator /= 2n;
  while (denominator % 5n === 0n) denominator /= 5n;
  if (denominator === 1n) return rendered;
  return value.denominator === '1'
    ? value.numerator
    : formatCellValue(value, { notation: 'fraction', approximationMarker: false });
}
function inlineHostOperands(
  node: FormulaNode,
  formula: string,
  sheetId: string,
  x: number,
  y: number,
  result: CalculationResult,
): string | null {
  const replacements: { start: number; end: number; text: string }[] = [];
  let unavailable = false;
  const visit = (node: FormulaNode) => {
    const host =
      node.kind === 'name' ||
      (node.kind === 'call' && /^(?:WORKITEM(?:\.|$)|DOCUMENT\.|USER\.)/.test(node.name));
    if (
      node.kind === 'quantity' ||
      (node.kind === 'call' && ['EXPM1', 'LOG1P'].includes(node.name))
    ) {
      unavailable = true;
      return;
    }
    if (host) {
      const stored = node.span
        ? result.context?.operands[formulaOperandKey(sheetId, x, y, node.span.start, node.span.end)]
        : undefined;
      if (
        !stored ||
        stored.formula !== '=' + formula ||
        !node.span ||
        (isNumeric(stored.value) && stored.value.dimension)
      ) {
        unavailable = true;
        return;
      }
      const value = stored.value;
      if (
        (typeof value === 'string' && value.length > 255) ||
        (isError(value) &&
          !['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'].includes(value.code))
      ) {
        unavailable = true;
        return;
      }
      const text = isNumeric(value)
        ? value.denominator === '1'
          ? value.numerator
          : `(${value.numerator}/${value.denominator})`
        : typeof value === 'string'
          ? '"' + value.replaceAll('"', '""') + '"'
          : value === null
            ? '""'
            : typeof value === 'boolean'
              ? value
                ? 'TRUE'
                : 'FALSE'
              : isError(value)
                ? value.code
                : '""';
      replacements.push({ ...node.span, text });
      return;
    }
    if (node.kind === 'call') {
      const aliases: Record<string, string> = {
        AVG: 'AVERAGE',
        CEILINGMATH: 'CEILING.MATH',
        CEILINGPRECISE: 'CEILING.PRECISE',
        FLOORMATH: 'FLOOR.MATH',
        FLOORPRECISE: 'FLOOR.PRECISE',
        RANKAVG: 'RANK.AVG',
        RANKEQ: 'RANK.EQ',
        STDEVS: 'STDEV.S',
        VARS: 'VAR.S',
      };
      if (node.span && Object.hasOwn(aliases, node.name))
        replacements.push({
          start: node.span.start,
          end: node.span.start + node.name.length,
          text: aliases[node.name],
        });
      node.args.forEach(visit);
    } else if (node.kind === 'binary') {
      visit(node.left);
      visit(node.right);
    } else if (node.kind === 'unary') visit(node.operand);
  };
  visit(node);
  if (unavailable) return null;
  let output = formula;
  for (const replacement of replacements.sort((a, b) => b.start - a.start))
    output = output.slice(0, replacement.start) + replacement.text + output.slice(replacement.end);
  return output.length > 8192 ? null : output;
}
function unsupportedCall(node: FormulaNode): boolean {
  if (node.kind === 'call')
    return (
      (!SUPPORTED_FUNCTIONS.includes(node.name) &&
        !/^(?:WORKITEM(?:\.|$)|DOCUMENT\.|USER\.)/.test(node.name)) ||
      node.args.some(unsupportedCall)
    );
  if (node.kind === 'binary') return unsupportedCall(node.left) || unsupportedCall(node.right);
  return node.kind === 'unary' ? unsupportedCall(node.operand) : false;
}
/** Exports a captured authored workbook and its existing results. It never evaluates formulas. */
export async function exportXlsx(
  workbook: Workbook,
  result: CalculationResult,
): Promise<Uint8Array> {
  assertWorkbook(workbook, { maxCells: 250000, maxWorksheets: 200 });
  if (result.schemaVersion !== 1 || result.worksheets.length !== workbook.worksheets.length)
    throw new TypeError('Export needs matching calculated worksheets.');
  const document = new ExcelJS.Workbook(),
    names = new Set<string>();
  document.calcProperties = {
    fullCalcOnLoad: false,
    forceFullCalc: false,
    calcMode: 'manual',
  } as ExcelJS.CalculationProperties;
  workbook.worksheets.forEach((sheet, index) => {
    if (
      !sheet.name ||
      sheet.name.length > 31 ||
      /[[\]:*?/\\]/.test(sheet.name) ||
      sheet.name.toLowerCase() === 'history' ||
      names.has(sheet.name.toLowerCase())
    )
      throw new TypeError('Worksheet names must follow Excel naming rules.');
    names.add(sheet.name.toLowerCase());
    const calculated = result.worksheets[index];
    if (
      calculated.id !== sheet.id ||
      calculated.values.length !== sheet.data.length ||
      calculated.values.some((row, y) => row.length !== sheet.data[y].length)
    )
      throw new TypeError('Export needs results matching every authored cell.');
    const output = document.addWorksheet(sheet.name);
    sheet.columns?.forEach((column, x) => {
      output.getColumn(x + 1).width = Math.max(1, ((column.width ?? 100) - 5) / 7);
    });
    for (const [row, geometry] of Object.entries(sheet.rows ?? {}))
      output.getRow(Number(row) + 1).height = (geometry.height * 3) / 4;
    sheet.data.forEach((row, y) =>
      row.forEach((input, x) => {
        const address = formatSpreadsheetAddress(x, y),
          cell = output.getCell(y + 1, x + 1),
          value = calculated.values[y][x],
          format = sheet.formats?.[address];
        const raw = typeof input === 'string' ? input : null;
        if (format?.kind === 'unit' || (isNumeric(value) && value.dimension))
          cell.value = formatWorksheetCell(value, format, { significantDigits: 10000 });
        else if (raw?.trimStart().startsWith('=')) {
          try {
            const formula = raw.trimStart().slice(1),
              node = parseFormula(formula);
            const inlined = inlineHostOperands(node, formula, sheet.id, x, y, result);
            if (
              inlined === null ||
              (isNumeric(value) && value.dimension) ||
              (isError(value) && ['#CYCLE!', '#LIMIT!', '#ERROR!'].includes(value.code))
            )
              cell.value = excelValue(value);
            else {
              const cache = excelValue(value);
              cell.value = {
                formula: inlined,
                result:
                  cache === null || unsupportedCall(node)
                    ? undefined
                    : (cache as ExcelJS.CellFormulaValue['result']),
              };
            }
          } catch {
            cell.value = raw;
          }
        } else if (typeof value === 'string') cell.value = value;
        else cell.value = excelValue(value);
        cell.style = {
          ...exportCellStyle(parseCellStyle(sheet.style?.[address])),
          numFmt: exportNumberFormat(format),
        };
      }),
    );
    for (const [address, [columns, rows]] of Object.entries(sheet.mergeCells ?? {})) {
      const anchor = parseSpreadsheetAddress(address)!;
      output.mergeCells(anchor.y + 1, anchor.x + 1, anchor.y + rows, anchor.x + columns);
    }
    if (sheet.freezeColumns)
      output.views = [{ state: 'frozen', xSplit: sheet.freezeColumns, ySplit: 0 }];
  });
  const files = unzipSync(new Uint8Array(await document.xlsx.writeBuffer()));
  const xml = strFromU8(files['xl/workbook.xml']);
  files['xl/workbook.xml'] = strToU8(
    xml.replace(
      /<calcPr\b[^>]*\/?\s*>/,
      '<calcPr calcMode="manual" fullCalcOnLoad="0" forceFullCalc="0" calcOnSave="0"/>',
    ),
  );
  return zipSync(files);
}
