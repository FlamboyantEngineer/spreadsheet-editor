import { parseCellStyle } from '../model/editing.js';
import type { WorksheetExtent } from '../model/extent.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetAddress,
} from '../model/spreadsheet-addressing.js';
import type { Worksheet } from '../model/workbook.js';
import {
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEADER_WIDTH,
  DEFAULT_ROW_HEIGHT,
  gridZoomScale,
  syncFrozenColumns,
} from './grid-geometry.js';
import {
  applyJSpreadsheetColumnMeasure,
  applyJSpreadsheetRowMeasure,
} from './jspreadsheet-measures.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import { restoreSharedGridBorders, syncSharedGridBorders } from './shared-grid-borders.js';
import { gridWorksheetExtent } from './worksheet-data.js';

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

type EngineOptions = NonNullable<JSpreadsheetWorksheetInstance['options']>;
/** The two versions of a worksheet a presentation pass moves between. */
interface WorksheetVersions {
  previous: Worksheet | null;
  next: Worksheet;
}
/** How the grid is viewed: editable or not, at a zoom, over the logical extent. */
interface WorksheetView {
  readOnly: boolean;
  zoom?: number | undefined;
  logicalExtent?: WorksheetExtent | undefined;
}

/** Writes every line's scaled size into the engine's options and measures. */
function syncLineSizes(
  engine: JSpreadsheetWorksheetInstance,
  options: EngineOptions,
  { next, scale, extent }: { next: Worksheet; scale: number; extent: WorksheetExtent | undefined },
) {
  const [columns, rows] = gridWorksheetExtent(next, extent);
  for (let x = 0; x < columns; x++) {
    const width = (next.columns?.[x]?.width ?? DEFAULT_COLUMN_WIDTH) * scale;
    options.columns ??= [];
    options.columns[x] = { ...options.columns[x], width, type: 'text' };
    applyJSpreadsheetColumnMeasure(engine, x, { attributeWidth: String(width) });
  }
  for (let y = 0; y < rows; y++) {
    const height = (next.rows?.[String(y)]?.height ?? DEFAULT_ROW_HEIGHT) * scale;
    (options.rows ??= {})[y] = { height };
    applyJSpreadsheetRowMeasure(engine, y, {
      cssHeight: `${height}px`,
      attributeHeight: String(height),
    });
  }
  options.defaultColWidth = DEFAULT_COLUMN_WIDTH * scale;
  options.defaultRowHeight = DEFAULT_ROW_HEIGHT * scale;
}

/** A pass over the engine's cells at one zoom scale. */
interface Pass {
  engine: JSpreadsheetWorksheetInstance;
  scale: number;
  zoomChanged: boolean;
}
const stale = (cell: HTMLElement, scale: number) => presentedScale.get(cell) !== scale;

/** Writes the authored visual style of one cell over its previous one. */
function restyleCell(cell: HTMLElement, before: string | undefined, after: string | undefined) {
  const old = parseCellStyle(before);
  const style = parseCellStyle(after);
  for (const property of new Set([...Object.keys(old), ...Object.keys(style)]))
    if (visualProperties.has(property)) cell.style.setProperty(property, style[property] ?? '');
}

/** Restyles the cells whose authored style changed, or all styled cells on a zoom change. */
function restyleCells(
  { engine, scale, zoomChanged }: Pass,
  { previous, next }: WorksheetVersions,
  pending: Set<HTMLElement>,
) {
  const addresses = new Set([
    ...Object.keys(previous?.style ?? {}),
    ...Object.keys(next.style ?? {}),
  ]);
  for (const address of addresses) {
    const coords = parseSpreadsheetAddress(address);
    const cell = coords && engine.getCellFromCoords?.(coords.x, coords.y);
    if (!cell || cell.classList.contains('editor')) continue;
    const before = previous?.style?.[address],
      after = next.style?.[address];
    if (!zoomChanged && !stale(cell, scale) && before === after) continue;
    restyleCell(cell, before, after);
    pending.add(cell);
  }
}

/** Adds the cells whose text metrics are stale, and cells deferred while they were edited. */
function staleCells({ engine, scale, zoomChanged }: Pass, pending: Set<HTMLElement>) {
  // Unscaled cells carry no presentation metrics, so only cells with authored styles
  // need the pass below; scaled cells need it once each, and all of them on a zoom change.
  if (zoomChanged || scale !== 1)
    for (const cell of engine.table?.querySelectorAll<HTMLElement>('thead td, tbody td') ?? [])
      if (zoomChanged || stale(cell, scale)) pending.add(cell);
  for (const cell of deferredCells.get(engine) ?? []) if (cell.isConnected) pending.add(cell);
}

/** Presents the attached pending cells; cells being edited wait for the next pass. */
function presentPending(pass: Pass, pending: Set<HTMLElement>, authored: Worksheet['style']) {
  const editing = new Set([...pending].filter((cell) => cell.classList.contains('editor')));
  deferredCells.set(pass.engine, editing);
  // Detached rows are presented when the row window attaches them (SE-PERF-1.1).
  const attached: HTMLElement[] = [];
  for (const cell of pending) {
    if (editing.has(cell)) continue;
    if (cell.isConnected) attached.push(cell);
    else presentedScale.delete(cell);
  }
  presentCells(attached, authored, pass.scale);
}

/** Scales the row-header column and the column-header row. */
function syncHeaderBands(engine: JSpreadsheetWorksheetInstance, scale: number) {
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
}

/** Removes merges that changed or went away, then sets the merges that are new or changed. */
function syncMerges(engine: JSpreadsheetWorksheetInstance, next: Worksheet) {
  const oldMerges = engine.getMerge?.();
  const known = oldMerges && !Array.isArray(oldMerges) ? oldMerges : undefined;
  for (const [address, span] of Object.entries(known ?? {}))
    if (JSON.stringify(span) !== JSON.stringify(next.mergeCells?.[address]))
      engine.removeMerge?.(address);
  for (const [address, span] of Object.entries(next.mergeCells ?? {}))
    if (!known || JSON.stringify(known[address]) !== JSON.stringify(span))
      engine.setMerge?.(address, span[0], span[1]);
}

/** Update the adopted engine without replacing its root, selection owner or event handlers.
 * Per-cell work is limited to cells that need it: every cell when zoom changes, otherwise
 * cells whose authored style changed and cells the engine has created since the last pass.
 */
export function syncWorksheetPresentation(
  engine: JSpreadsheetWorksheetInstance,
  versions: WorksheetVersions,
  view: WorksheetView,
) {
  const options = engine.options;
  if (!options) return;
  const { next } = versions;
  restoreSharedGridBorders(engine);
  options.editable = !view.readOnly;
  const scale = gridZoomScale(view.zoom ?? 100);
  const zoomChanged = engineScale.get(engine) !== scale;
  engineScale.set(engine, scale);
  syncLineSizes(engine, options, { next, scale, extent: view.logicalExtent });
  const pass: Pass = { engine, scale, zoomChanged };
  const pending = new Set<HTMLElement>();
  restyleCells(pass, versions, pending);
  staleCells(pass, pending);
  presentPending(pass, pending, next.style);
  syncHeaderBands(engine, scale);
  // The engine's `options.style` is a getter that reads every cell's style attribute.
  authoredStyles.set(engine, next.style);
  syncMerges(engine, next);
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
