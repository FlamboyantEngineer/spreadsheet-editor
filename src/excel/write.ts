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
import { exportCellStyle, exportNumberFormat, numberFormatLoss } from './styles.js';
import { assertCalculationResult } from '../calculation/http.js';
import type {
  ExcelFormulaCapabilities,
  ExcelHostFunction,
  XlsxExport,
  XlsxWarning,
} from './contracts.js';
import { hostFunctionMap, validateFormulaCapabilities } from './capabilities.js';
import { capturedText } from './captured-text.js';

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
  if (denominator === 1n && exact.compare(exact.parse(rendered), value) === 0) return rendered;
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
  functions: ReadonlyMap<string, ExcelHostFunction>,
  loss: (code: string, message: string) => void,
): string | null {
  const replacements: { start: number; end: number; text: string }[] = [];
  let unavailable = false;
  const visit = (node: FormulaNode) => {
    const host = node.kind === 'name' || (node.kind === 'call' && functions.has(node.name));
    if (
      node.kind === 'quantity' ||
      (node.kind === 'call' && ['EXPM1', 'LOG1P'].includes(node.name))
    ) {
      unavailable = true;
      loss(
        'formula-stored-value',
        'Excel has no equivalent for this quantity or numerical extension; the stored result is exported.',
      );
      return;
    }
    if (node.kind === 'literal' && typeof node.value === 'string' && node.value.length > 255) {
      unavailable = true;
      loss(
        'formula-string-limit',
        'A formula string exceeds Excel’s 255-character limit; the stored result is exported.',
      );
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
        (isNumeric(stored.value) && (stored.value.dimension || stored.value.quantityKind))
      ) {
        unavailable = true;
        loss(
          'captured-operand',
          'A matching representable captured operand is unavailable; the stored result is exported.',
        );
        return;
      }
      let value = stored.value;
      if (typeof value === 'string') {
        try {
          value = capturedText(
            value,
            node.kind === 'call' ? (functions.get(node.name)?.capturedText ?? 'plain') : 'plain',
          );
        } catch (error) {
          unavailable = true;
          loss(
            'captured-text',
            (error instanceof Error ? error.message : 'Captured text cannot be represented.') +
              ' The stored result is exported.',
          );
          return;
        }
      }
      if (
        (typeof value === 'string' && value.length > 255) ||
        (isError(value) &&
          !['#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'].includes(value.code))
      ) {
        unavailable = true;
        loss(
          'captured-operand',
          'Excel cannot represent the captured operand; the stored result is exported.',
        );
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
  if (output.length > 8192) {
    loss(
      'formula-length',
      'The expanded formula exceeds Excel’s 8192-character limit; the stored result is exported.',
    );
    return null;
  }
  return output;
}
function unsupportedCall(
  node: FormulaNode,
  functions: ReadonlyMap<string, ExcelHostFunction>,
): boolean {
  if (node.kind === 'call')
    return (
      (!SUPPORTED_FUNCTIONS.includes(node.name) && !functions.has(node.name)) ||
      (!functions.has(node.name) && node.args.some((arg) => unsupportedCall(arg, functions)))
    );
  if (node.kind === 'binary')
    return unsupportedCall(node.left, functions) || unsupportedCall(node.right, functions);
  return node.kind === 'unary' ? unsupportedCall(node.operand, functions) : false;
}
/** Exports a captured authored workbook and its existing results. It never evaluates formulas. */
export async function exportXlsx(
  workbook: Workbook,
  result: CalculationResult,
  options: { formulaCapabilities?: ExcelFormulaCapabilities } = {},
): Promise<Uint8Array> {
  return (await exportXlsxWithReport(workbook, result, options)).bytes;
}

/** Captured-result export with explicit grouped losses; no live resolver is available. */
export async function exportXlsxWithReport(
  workbook: Workbook,
  result: CalculationResult,
  options: { formulaCapabilities?: ExcelFormulaCapabilities } = {},
): Promise<XlsxExport> {
  const functions = hostFunctionMap(validateFormulaCapabilities(options.formulaCapabilities));
  const warnings = new Map<string, XlsxWarning>();
  const warn = (
    code: string,
    message: string,
    location: string,
    disposition: XlsxWarning['disposition'] = 'stored-value',
  ) => {
    const entry = warnings.get(code) ?? { code, message, disposition, count: 0, locations: [] };
    entry.count++;
    if (entry.locations.length < 5 && !entry.locations.includes(location))
      entry.locations.push(location);
    warnings.set(code, entry);
  };
  const report = {
    worksheets: workbook.worksheets.length,
    cells: 0,
    formulas: 0,
    mergedRanges: 0,
    warnings: [] as XlsxWarning[],
  };
  assertWorkbook(workbook, { maxCells: 250000, maxWorksheets: 200 });
  assertCalculationResult(result, workbook);
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
          format = sheet.formats?.[address],
          location = `${sheet.name}!${address}`;
        report.cells++;
        const formattingLoss = numberFormatLoss(format);
        if (formattingLoss) warn('number-format', formattingLoss, location, 'not-exported');
        const canonical = (sheet as typeof sheet & { inputValues?: Record<string, unknown> })
          .inputValues?.[address];
        if (canonical)
          warn(
            'canonical-input',
            'Canonical numeric metadata is projected as the stored value; XLSX does not retain the canonical override or its provenance.',
            location,
            'not-exported',
          );
        if (
          isNumeric(value) &&
          ('temperature' in value || value.quantityKind || 'provenance' in value)
        )
          warn(
            'quantity-metadata',
            'Quantity and provenance metadata are retained in displayed stored text where applicable, but XLSX does not retain the typed metadata.',
            location,
            'not-exported',
          );
        const raw = typeof input === 'string' ? input : null;
        if (raw?.trimStart().startsWith('=')) report.formulas++;
        if (
          format?.kind === 'unit' ||
          (isNumeric(value) && (value.dimension || value.quantityKind))
        ) {
          cell.value =
            format?.kind === 'unit'
              ? formatWorksheetCell(value, format, { significantDigits: 10000 })
              : formatCellValue(value, { significantDigits: 10000 });
          warn(
            'quantity-text',
            'A quantity is exported as stored text to preserve its units and temperature meaning.',
            location,
          );
        } else if (raw?.trimStart().startsWith('=')) {
          try {
            const formula = raw.trimStart().slice(1),
              node = parseFormula(formula);
            const inlined = inlineHostOperands(
              node,
              formula,
              sheet.id,
              x,
              y,
              result,
              functions,
              (code, message) => warn(code, message, location),
            );
            if (
              inlined === null ||
              (isNumeric(value) && value.dimension) ||
              (isError(value) && ['#CYCLE!', '#LIMIT!', '#ERROR!'].includes(value.code))
            ) {
              cell.value = excelValue(value);
              if (isError(value) && ['#CYCLE!', '#LIMIT!', '#ERROR!'].includes(value.code))
                warn(
                  'formula-error',
                  'Excel cannot represent this calculation error; its stored error text is exported.',
                  location,
                );
            } else {
              const cache = excelValue(value),
                unsupported = unsupportedCall(node, functions);
              if (unsupported)
                warn(
                  'unsupported-formula',
                  'An unsupported or undeclared host formula is retained without a cache.',
                  location,
                  'not-exported',
                );
              cell.value = {
                formula: inlined,
                result:
                  cache === null || unsupported
                    ? undefined
                    : (cache as ExcelJS.CellFormulaValue['result']),
              };
            }
          } catch {
            warn(
              'malformed-formula',
              'The formula is exported as authored text because it cannot be parsed.',
              location,
              'not-exported',
            );
            cell.value = raw;
          }
        } else if (typeof value === 'string') cell.value = value;
        else cell.value = excelValue(value);
        if (typeof cell.value === 'string' && cell.value.length > 32767)
          throw new RangeError(`${location}: Excel cell strings have a 32767-character limit.`);
        cell.style = {
          ...exportCellStyle(parseCellStyle(sheet.style?.[address])),
          numFmt: exportNumberFormat(format),
        };
      }),
    );
    for (const [address, [columns, rows]] of Object.entries(sheet.mergeCells ?? {})) {
      report.mergedRanges++;
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
  report.warnings = [...warnings.values()];
  return { bytes: zipSync(files), report };
}
