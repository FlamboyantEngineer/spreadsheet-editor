import type { Worksheet } from '../model/workbook.js';
import type { WorksheetExtent } from '../model/extent.js';
import { parseCellStyle } from '../model/editing.js';
import { gridWorksheetExtent } from './worksheet-data.js';
import { restoreSharedGridBorders, syncSharedGridBorders } from './shared-grid-borders.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetAddress,
} from '../model/spreadsheet-addressing.js';
import {
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  DEFAULT_ROW_HEADER_WIDTH,
  gridZoomScale,
  syncFrozenColumns,
} from './grid-geometry.js';
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
/** The zoom scale each cell's inline text metrics were last written for. */
const presentedScale = new WeakMap<HTMLElement, number>();
const engineScale = new WeakMap<JSpreadsheetWorksheetInstance, number>();
const authoredStyles = new WeakMap<JSpreadsheetWorksheetInstance, Worksheet['style']>();
/** Cells being edited during a zoom change; the next pass presents them. */
const deferredCells = new WeakMap<JSpreadsheetWorksheetInstance, Set<HTMLElement>>();

/** Update the adopted engine without replacing its root, selection owner or event handlers.
 * Per-cell work is limited to cells that need it: every cell when zoom changes, otherwise
 * cells whose authored style changed and cells the engine has created since the last pass.
 */
export function syncWorksheetPresentation(
  engine: JSpreadsheetWorksheetInstance,
  previous: Worksheet | null,
  next: Worksheet,
  readOnly: boolean,
  zoom = 100,
  logicalExtent?: WorksheetExtent,
) {
  const options = engine.options;
  if (!options) return;
  restoreSharedGridBorders(engine);
  options.editable = !readOnly;
  const scale = gridZoomScale(zoom);
  const zoomChanged = engineScale.get(engine) !== scale;
  engineScale.set(engine, scale);
  const [columns, rows] = gridWorksheetExtent(next, logicalExtent);
  for (let x = 0; x < columns; x++) {
    const width = (next.columns?.[x]?.width ?? DEFAULT_COLUMN_WIDTH) * scale;
    options.columns ??= [];
    options.columns[x] = { ...options.columns[x], width, type: 'text' };
    applyJSpreadsheetColumnMeasure(engine, x, String(width));
  }
  for (let y = 0; y < rows; y++) {
    const height = (next.rows?.[String(y)]?.height ?? DEFAULT_ROW_HEIGHT) * scale;
    (options.rows ??= {})[y] = { height };
    applyJSpreadsheetRowMeasure(engine, y, `${height}px`, String(height));
  }
  options.defaultColWidth = DEFAULT_COLUMN_WIDTH * scale;
  options.defaultRowHeight = DEFAULT_ROW_HEIGHT * scale;
  const stale = (cell: HTMLElement) => presentedScale.get(cell) !== scale;
  const pending = new Set<HTMLElement>();
  const addresses = new Set([
    ...Object.keys(previous?.style ?? {}),
    ...Object.keys(next.style ?? {}),
  ]);
  for (const address of addresses) {
    const coords = parseSpreadsheetAddress(address);
    const cell = coords && engine.getCellFromCoords?.(coords.x, coords.y);
    if (!cell || cell.classList.contains('editor')) continue;
    if (!zoomChanged && !stale(cell) && previous?.style?.[address] === next.style?.[address])
      continue;
    const old = parseCellStyle(previous?.style?.[address]);
    const style = parseCellStyle(next.style?.[address]);
    for (const property of new Set([...Object.keys(old), ...Object.keys(style)])) {
      if (visualProperties.has(property)) cell.style.setProperty(property, style[property] ?? '');
    }
    pending.add(cell);
  }
  // Unscaled cells carry no presentation metrics, so only cells with authored styles
  // need the pass below; scaled cells need it once each, and all of them on a zoom change.
  if (zoomChanged || scale !== 1)
    for (const cell of engine.table?.querySelectorAll<HTMLElement>('thead td, tbody td') ?? [])
      if (zoomChanged || stale(cell)) pending.add(cell);
  for (const cell of deferredCells.get(engine) ?? []) if (cell.isConnected) pending.add(cell);
  const editing = new Set([...pending].filter((cell) => cell.classList.contains('editor')));
  deferredCells.set(engine, editing);
  // Detached rows are presented when the row window attaches them (SE-PERF-1).
  const attached: HTMLElement[] = [];
  for (const cell of pending) {
    if (editing.has(cell)) continue;
    if (cell.isConnected) attached.push(cell);
    else presentedScale.delete(cell);
  }
  presentCells(attached, next.style, scale);
  const table = engine.table;
  const cornerColumn = table?.querySelector<HTMLElement>('colgroup > col:first-child');
  if (cornerColumn) {
    const width = DEFAULT_ROW_HEADER_WIDTH * scale;
    if (cornerColumn.style.width !== `${width}px`) cornerColumn.style.width = `${width}px`;
    if (cornerColumn.getAttribute('width') !== String(width))
      cornerColumn.setAttribute('width', String(width));
  }
  const header = table?.tHead?.rows[0];
  const headerHeight = `${DEFAULT_ROW_HEIGHT * scale}px`;
  if (header && header.style.height !== headerHeight) header.style.height = headerHeight;
  // The engine's `options.style` is a getter that reads every cell's style attribute.
  authoredStyles.set(engine, next.style);
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
  syncFrozenColumns(engine);
  syncSharedGridBorders(engine);
}

/** Write zoomed text metrics on cells that are attached, from their computed style. */
function presentCells(cells: HTMLElement[], authored: Worksheet['style'], scale: number) {
  // Resolve font units through CSS after restoring authored styles. Presentation-only
  // inline values are reset first so successive zooms never compound their scale.
  for (const cell of cells) {
    const x = Number(cell.dataset.x),
      y = Number(cell.dataset.y);
    const address =
      Number.isInteger(x) && Number.isInteger(y) && x >= 0 && y >= 0
        ? formatSpreadsheetAddress(x, y)
        : undefined;
    const style = parseCellStyle(address ? authored?.[address] : undefined);
    cell.style.fontSize = style['font-size'] ?? '';
    cell.style.lineHeight = '';
    cell.style.padding = '';
    cell.style.textIndent = style['text-indent'] ?? '';
  }
  const computed = cells.map((cell) => {
    const current = getComputedStyle(cell);
    return {
      cell,
      fontSize: Number.parseFloat(current.fontSize),
      lineHeight: Number.parseFloat(current.lineHeight),
      padding: [
        current.paddingTop,
        current.paddingRight,
        current.paddingBottom,
        current.paddingLeft,
      ].map((size) => Number.parseFloat(size) || 0),
      indent: Number.parseFloat(current.textIndent) || 0,
      wideBorders: (['Top', 'Right', 'Bottom', 'Left'] as const).filter(
        (side) =>
          current[`border${side}Style`] !== 'none' &&
          ![0, 1].includes(Number.parseFloat(current[`border${side}Width`])),
      ),
    };
  });
  for (const { cell, fontSize, lineHeight, padding, indent, wideBorders } of computed) {
    // Authored border colors/styles remain intact while stroke widths stay one CSS pixel.
    for (const side of wideBorders) cell.style[`border${side}Width`] = '1px';
    presentedScale.set(cell, scale);
    if (scale === 1) continue;
    if (fontSize) cell.style.fontSize = `${fontSize * scale}px`;
    if (lineHeight) cell.style.lineHeight = `${lineHeight * scale}px`;
    cell.style.padding = padding.map((size) => `${size * scale}px`).join(' ');
    cell.style.textIndent = `${indent * scale}px`;
  }
}

/**
 * Present the cells of rows the row window has just attached: cells written for another zoom,
 * and, away from 100%, cells never presented. At 100% a cell without inline metrics needs
 * nothing unless it has an authored style.
 */
export function presentAttachedRows(engine: JSpreadsheetWorksheetInstance, rows: number[]) {
  const scale = engineScale.get(engine);
  if (scale === undefined) return;
  const authored = authoredStyles.get(engine);
  const cells: HTMLElement[] = [];
  for (const y of rows)
    for (const record of engine.records?.[y] ?? []) {
      const cell = record?.element as HTMLElement | undefined;
      if (!cell || cell.classList.contains('editor')) continue;
      const presented = presentedScale.get(cell);
      if (
        presented === undefined
          ? scale !== 1 || !!authored?.[formatSpreadsheetAddress(Number(cell.dataset.x), y)]
          : presented !== scale
      )
        cells.push(cell);
    }
  if (cells.length) presentCells(cells, authored, scale);
}
