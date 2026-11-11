import type { Worksheet } from '../model/workbook.js';
import { parseCellStyle } from '../model/editing.js';
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
/** Update the adopted engine without replacing its root, selection owner or event handlers. */
export function syncWorksheetPresentation(
  engine: JSpreadsheetWorksheetInstance,
  previous: Worksheet | null,
  next: Worksheet,
  readOnly: boolean,
  zoom = 100,
) {
  const options = engine.options;
  if (!options) return;
  restoreSharedGridBorders(engine);
  options.editable = !readOnly;
  const scale = gridZoomScale(zoom);
  const columns = Math.max(
    engine.cols?.length ?? 0,
    next.minDimensions?.[0] ?? 26,
    previous?.columns?.length ?? 0,
    next.columns?.length ?? 0,
  );
  for (let x = 0; x < columns; x++) {
    const width = (next.columns?.[x]?.width ?? DEFAULT_COLUMN_WIDTH) * scale;
    options.columns ??= [];
    options.columns[x] = { ...options.columns[x], width, type: 'text' };
    applyJSpreadsheetColumnMeasure(engine, x, String(width));
  }
  const rows = Math.max(
    engine.rows?.length ?? 0,
    next.minDimensions?.[1] ?? 100,
    next.data.length,
    ...Object.keys(next.rows ?? {}).map((key) => Number(key) + 1),
  );
  for (let y = 0; y < rows; y++) {
    const height = (next.rows?.[String(y)]?.height ?? DEFAULT_ROW_HEIGHT) * scale;
    (options.rows ??= {})[y] = { height };
    applyJSpreadsheetRowMeasure(engine, y, `${height}px`, String(height));
  }
  options.defaultColWidth = DEFAULT_COLUMN_WIDTH * scale;
  options.defaultRowHeight = DEFAULT_ROW_HEIGHT * scale;
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
  // Resolve font units through CSS after restoring authored styles. Presentation-only
  // inline values are reset first so successive zooms never compound their scale.
  const cells = Array.from(
    engine.table?.querySelectorAll<HTMLElement>('thead td, tbody td') ?? [],
  ).filter((cell) => !cell.classList.contains('editor'));
  for (const cell of cells) {
    const x = Number(cell.dataset.x),
      y = Number(cell.dataset.y);
    const address =
      Number.isInteger(x) && Number.isInteger(y) && x >= 0 && y >= 0
        ? formatSpreadsheetAddress(x, y)
        : undefined;
    const style = parseCellStyle(address ? next.style?.[address] : undefined);
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
      borders: (['Top', 'Right', 'Bottom', 'Left'] as const).filter(
        (side) =>
          current[`border${side}Style`] !== 'none' &&
          Number.parseFloat(current[`border${side}Width`]) > 0,
      ),
    };
  });
  for (const { cell, fontSize, lineHeight, padding, indent, borders } of computed) {
    // Authored border colors/styles remain intact while stroke widths stay one CSS pixel.
    for (const side of borders) cell.style[`border${side}Width`] = '1px';
    if (fontSize) cell.style.fontSize = `${fontSize * scale}px`;
    if (lineHeight) cell.style.lineHeight = `${lineHeight * scale}px`;
    cell.style.padding = padding.map((size) => `${size * scale}px`).join(' ');
    cell.style.textIndent = `${indent * scale}px`;
  }
  const table = engine.table;
  const cornerColumn = table?.querySelector<HTMLElement>('colgroup > col:first-child');
  if (cornerColumn) {
    cornerColumn.style.width = `${DEFAULT_ROW_HEADER_WIDTH * scale}px`;
    cornerColumn.setAttribute('width', String(DEFAULT_ROW_HEADER_WIDTH * scale));
  }
  const header = table?.tHead?.rows[0];
  if (header) header.style.height = `${DEFAULT_ROW_HEIGHT * scale}px`;
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
  syncFrozenColumns(engine);
  syncSharedGridBorders(engine);
}
