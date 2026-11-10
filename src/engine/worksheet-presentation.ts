import type { Worksheet } from '../model/workbook.js';
import { parseCellStyle } from '../model/editing.js';
import { parseSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import {
  applyJSpreadsheetColumnMeasure,
  applyJSpreadsheetRowMeasure,
  type JSpreadsheetWorksheetInstance,
} from './jspreadsheet-host.js';

const visualProperties = new Set([
  'font-family',
  'font-size',
  'font-weight',
  'font-style',
  'text-decoration',
  'color',
  'background-color',
  'text-align',
  'vertical-align',
  'white-space',
  'border',
  'border-top',
  'border-right',
  'border-bottom',
  'border-left',
  'text-indent',
]);
/** Update the adopted engine without replacing its root, selection owner or event handlers. */
export function syncWorksheetPresentation(
  engine: JSpreadsheetWorksheetInstance,
  previous: Worksheet | null,
  next: Worksheet,
  readOnly: boolean,
) {
  const options = engine.options;
  if (!options) return;
  options.editable = !readOnly;
  const columns = Math.max(previous?.columns?.length ?? 0, next.columns?.length ?? 0);
  for (let x = 0; x < columns; x++) {
    const width = next.columns?.[x]?.width ?? 100;
    options.columns ??= [];
    options.columns[x] = { ...options.columns[x], width, type: 'text' };
    applyJSpreadsheetColumnMeasure(engine, x, String(width));
  }
  for (const key of new Set([
    ...Object.keys(previous?.rows ?? {}),
    ...Object.keys(next.rows ?? {}),
  ])) {
    const height = next.rows?.[key]?.height ?? 24;
    (options.rows ??= {})[Number(key)] = { height };
    applyJSpreadsheetRowMeasure(engine, Number(key), `${height}px`, String(height));
  }
  const addresses = new Set([
    ...Object.keys(previous?.style ?? {}),
    ...Object.keys(next.style ?? {}),
  ]);
  for (const address of addresses) {
    const coords = parseSpreadsheetAddress(address);
    const cell = coords && engine.getCellFromCoords?.(coords.x, coords.y);
    if (!cell || cell.classList.contains('editor')) continue;
    const old = parseCellStyle(previous?.style?.[address]);
    const style = parseCellStyle(next.style?.[address]);
    for (const property of new Set([...Object.keys(old), ...Object.keys(style)])) {
      if (visualProperties.has(property)) cell.style.setProperty(property, style[property] ?? '');
    }
  }
  const storedStyle = options.style;
  if (storedStyle) {
    for (const address of Object.keys(storedStyle)) delete storedStyle[address];
    Object.assign(storedStyle, next.style);
  }
  const oldMerges = engine.getMerge?.();
  if (oldMerges && !Array.isArray(oldMerges)) {
    for (const [address, span] of Object.entries(oldMerges))
      if (JSON.stringify(span) !== JSON.stringify(next.mergeCells?.[address]))
        engine.removeMerge?.(address);
  }
  for (const [address, span] of Object.entries(next.mergeCells ?? {})) {
    if (
      !oldMerges ||
      Array.isArray(oldMerges) ||
      JSON.stringify(oldMerges[address]) !== JSON.stringify(span)
    )
      engine.setMerge?.(address, span[0], span[1]);
  }
  options.freezeColumns = next.freezeColumns ?? 0;
  // CE exposes this method at runtime but omits it from its declarations.
  (
    engine as JSpreadsheetWorksheetInstance & { updateFreezePosition?: () => void }
  ).updateFreezePosition?.();
}
