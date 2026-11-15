// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import {
  clipGridSelectionHandle,
  expandMergedSelection,
  gridBodyBounds,
  gridZoomScale,
  restoreGridScroll,
  scrollGridCellIntoView,
  syncFrozenColumnOffsets,
  syncFrozenColumns,
} from './grid-geometry.js';
import { measureGridSelection } from './grid-selection.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';
import { syncWorksheetPresentation } from './worksheet-presentation.js';

function rect(left: number, top: number, width: number, height: number): DOMRect {
  return {
    left,
    top,
    right: left + width,
    bottom: top + height,
    width,
    height,
    x: left,
    y: top,
    toJSON: () => ({}),
  };
}
function grid() {
  const table = document.createElement('table');
  table.innerHTML =
    '<colgroup><col width="50"><col width="100"><col width="100"><col width="100"></colgroup><thead><tr><td></td><td data-x="0">A</td><td data-x="1">B</td><td data-x="2">C</td></tr></thead><tbody><tr><td class="jss_row">1</td><td data-x="0" data-y="0">1</td><td data-x="1" data-y="0">2</td><td data-x="2" data-y="0">3</td></tr><tr><td class="jss_row">2</td><td data-x="0" data-y="1">4</td><td data-x="1" data-y="1">5</td><td data-x="2" data-y="1">6</td></tr></tbody>';
  const viewport = document.createElement('div');
  viewport.append(table);
  document.body.append(viewport);
  const css = document.createElement('style');
  css.textContent =
    'table td { font-size:12px; padding:4px; line-height:12px; border:1px solid black; }';
  document.head.append(css);
  Object.defineProperties(viewport, { clientWidth: { value: 400 }, clientHeight: { value: 400 } });
  vi.spyOn(viewport, 'getBoundingClientRect').mockReturnValue(rect(10, 20, 400, 400));
  const cols = Array.from(table.querySelectorAll('col'))
    .slice(1)
    .map((colElement) => ({ colElement }));
  const rows = Array.from(table.tBodies[0].rows).map((element) => ({ element }));
  const records = rows.map((row) =>
    Array.from(row.element.cells)
      .slice(1)
      .map((element) => ({ element })),
  );
  const headers = Array.from(requireDefined(table.tHead).rows[0].cells).slice(1);
  vi.spyOn(requireDefined(table.tHead).rows[0].cells[0], 'getBoundingClientRect').mockReturnValue(
    rect(10, 20, 50, 24),
  );
  headers.forEach((header, x) =>
    vi.spyOn(header, 'getBoundingClientRect').mockReturnValue(rect(60 + x * 100, 20, 100, 24)),
  );
  records.forEach((row, y) =>
    row.forEach((cell, x) =>
      vi
        .spyOn(cell.element, 'getBoundingClientRect')
        .mockReturnValue(rect(60 + x * 100, 44 + y * 24, 100, 24)),
    ),
  );
  const engine = {
    table,
    content: viewport,
    cols,
    rows,
    records,
    headers,
    options: { columns: [], rows: {}, style: {}, freezeColumns: 0 },
    getCellFromCoords: (x: number, y: number) => records[y]?.[x]?.element ?? null,
  } as unknown as JSpreadsheetWorksheetInstance;
  const sheet = createWorkbook([
    ['1', '2', '3'],
    ['4', '5', '6'],
  ]).worksheets[0];
  return { engine, sheet, viewport, table, records, cols };
}
afterEach(() => {
  document.body.innerHTML = '';
  document.head.innerHTML = '';
  vi.restoreAllMocks();
});

describe('SE-CHROME-5.1: grid-only zoom keeps authored geometry and pixel strokes', () => {
  it('scales explicit and default measures, typography and headers without compounding', () => {
    const { engine, sheet, table, records, cols } = grid();
    sheet.columns = [{ width: 120.5 }];
    sheet.rows = { '0': { height: 35.5 } };
    sheet.style = { A1: 'font-size:16px;border-right:3px solid red' };
    const original = structuredClone(sheet);
    for (const zoom of [50, 75, 100, 125, 150, 100]) {
      syncWorksheetPresentation(
        engine,
        { previous: sheet, next: sheet },
        { readOnly: false, zoom },
      );
      expect(cols[0].colElement.style.width).toBe(`${(120.5 * zoom) / 100}px`);
      expect(cols[1].colElement.style.width).toBe(`${zoom}px`);
      expect(engine.rows?.[0].element.style.height).toBe(`${(35.5 * zoom) / 100}px`);
      expect(engine.rows?.[1].element.style.height).toBe(`${(24 * zoom) / 100}px`);
      expect(engine.options?.columns?.[25].width).toBe(zoom);
      expect(engine.options?.rows?.[99].height).toBe((24 * zoom) / 100);
      expect(requireDefined(table.querySelector<HTMLElement>('col')).style.width).toBe(
        `${(50 * zoom) / 100}px`,
      );
      // At 100% the stylesheet's own metrics apply; no inline copy is written.
      expect(getComputedStyle(records[0][0].element).fontSize).toBe(`${(16 * zoom) / 100}px`);
      expect(getComputedStyle(records[0][1].element).fontSize).toBe(`${(12 * zoom) / 100}px`);
      expect(records[0][1].element.style.fontSize).toBe(
        zoom === 100 ? '' : `${(12 * zoom) / 100}px`,
      );
      expect(records[0][0].element.style.borderRightWidth).toBe('1px');
      expect(sheet).toEqual(original);
      // The engine's style option is a getter over the cells' style attributes; it is not written.
    }
  });

  it('preserves the logical scroll anchor using only the inner viewport', () => {
    const { viewport } = grid();
    restoreGridScroll(viewport, { left: 150.5, top: 240.25 }, { from: 1, to: 1.5 });
    expect(viewport.scrollLeft).toBe(225.75);
    expect(viewport.scrollTop).toBe(360.375);
    expect(gridZoomScale(undefined)).toBe(1);
    expect(gridZoomScale(Number.NaN)).toBe(1);
  });

  it('uses fractional frozen offsets instead of truncating widths', () => {
    const { engine, cols, records } = grid();
    requireDefined(engine.options).freezeColumns = 2;
    cols[0].colElement.style.width = '150.75px';
    syncFrozenColumns(engine);
    expect(engine.headers?.[0].style.left).toBe('50px');
    expect(engine.headers?.[1].style.left).toBe('200.75px');
    expect(records[0][1].element.style.left).toBe('200.75px');
    expect(records[0][1].element.style.position).toBe('sticky');
    requireDefined(engine.options).freezeColumns = 0;
    syncFrozenColumns(engine);
    expect(records[0][1].element.style.left).toBe('');
    expect(records[0][1].element.classList.contains('jss_freezed')).toBe(false);
  });
  it('restores frozen scroll offsets without changing unfrozen cells or their classes', () => {
    const { engine, cols, records } = grid();
    requireDefined(engine.options).freezeColumns = 1;
    cols[0].colElement.style.width = '150.75px';
    syncFrozenColumns(engine);
    requireDefined(engine.headers)[0].style.left = '0px';
    records[0][0].element.style.left = '149px';
    const plain = records[0][1].element.getAttribute('style');
    syncFrozenColumnOffsets(engine);
    expect(requireDefined(engine.headers)[0].style.left).toBe('50px');
    expect(records[0][0].element.style.left).toBe('50px');
    expect(records[0][1].element.getAttribute('style')).toBe(plain);
  });
});

describe('SE-UI-1.1: local selection geometry', () => {
  it('clips the native fill handle to the cell body and restores it after scrolling back', () => {
    const { engine, viewport } = grid();
    const handle = document.createElement('div');
    engine.corner = handle;
    const measure = vi.spyOn(handle, 'getBoundingClientRect').mockReturnValue(rect(58, 40, 7, 7));
    const body = gridBodyBounds(engine, viewport);
    clipGridSelectionHandle(engine, body);
    expect(handle.style.clipPath).toBe('inset(4px 0px 0px 2px)');
    measure.mockReturnValue(rect(20, 30, 7, 7));
    clipGridSelectionHandle(engine, body);
    expect(handle.style.clipPath).toBe('inset(50%)');
    measure.mockReturnValue(rect(60, 44, 7, 7));
    clipGridSelectionHandle(engine, body);
    expect(handle.style.clipPath).toBe('');
  });
  it('clips to live row and column headers when the corner has scrolled away', () => {
    const { engine, sheet, viewport, table } = grid();
    vi.mocked(requireDefined(table.tHead).rows[0].cells[0].getBoundingClientRect).mockReturnValue(
      rect(-190, -80, 50, 24),
    );
    vi.spyOn(table.tBodies[0].rows[0].cells[0], 'getBoundingClientRect').mockReturnValue(
      rect(10, -20, 55, 24),
    );
    expect(gridBodyBounds(engine, viewport)).toEqual({
      origin: { left: 10, top: 20 },
      left: 65,
      top: 44,
      right: 410,
      bottom: 420,
    });
    expect(
      measureGridSelection(engine, { sheet, selection: { x1: 0, y1: 0, x2: 0, y2: 0 } }, viewport),
    ).toEqual([
      expect.objectContaining({ left: 55, top: 24, width: 95, height: 24, borderLeftWidth: 0 }),
    ]);
  });
  it('measures only rectangle corners, including with unrelated merges elsewhere', () => {
    const { engine, sheet, viewport, records } = grid();
    sheet.mergeCells = { Z99: [2, 2] };
    measureGridSelection(engine, { sheet, selection: { x1: 0, y1: 0, x2: 2, y2: 1 } }, viewport);
    const reads = records
      .flat()
      .reduce(
        (total, cell) => total + vi.mocked(cell.element.getBoundingClientRect).mock.calls.length,
        0,
      );
    expect(reads).toBe(4);
  });
  it('measures all four edges of a cell or reversed range in CSS pixels', () => {
    const { engine, sheet, viewport } = grid();
    expect(
      measureGridSelection(engine, { sheet, selection: { x1: 0, y1: 0, x2: 0, y2: 0 } }, viewport),
    ).toEqual([
      expect.objectContaining({
        left: 50,
        top: 24,
        width: 100,
        height: 24,
        position: 'absolute',
        pointerEvents: 'none',
        boxSizing: 'border-box',
      }),
    ]);
    expect(
      measureGridSelection(engine, { sheet, selection: { x1: 2, y1: 1, x2: 0, y2: 0 } }, viewport),
    ).toEqual([expect.objectContaining({ left: 50, top: 24, width: 300, height: 48 })]);
  });

  it('expands hidden merged endpoints to the full merged rectangle', () => {
    const { engine, sheet, viewport, records } = grid();
    sheet.mergeCells = { A1: [2, 2] };
    records[0][1].element.style.display = 'none';
    records[1][0].element.style.display = 'none';
    records[1][1].element.style.display = 'none';
    vi.mocked(records[0][0].element.getBoundingClientRect).mockReturnValue(rect(60, 44, 200, 48));
    expect(expandMergedSelection({ x1: 1, y1: 1, x2: 1, y2: 1 }, sheet)).toEqual({
      x1: 0,
      y1: 0,
      x2: 1,
      y2: 1,
    });
    expect(
      measureGridSelection(engine, { sheet, selection: { x1: 1, y1: 1, x2: 1, y2: 1 } }, viewport),
    ).toEqual([expect.objectContaining({ left: 50, top: 24, width: 200, height: 48 })]);
  });

  it('splits across frozen columns and hides clipped or internal stroke sides', () => {
    const { engine, sheet, viewport, records } = grid();
    sheet.freezeColumns = 1;
    vi.mocked(records[0][1].element.getBoundingClientRect).mockReturnValue(rect(110, 44, 100, 24));
    vi.mocked(records[0][2].element.getBoundingClientRect).mockReturnValue(rect(210, 44, 100, 24));
    const pieces = measureGridSelection(
      engine,
      { sheet, selection: { x1: 0, y1: 0, x2: 2, y2: 0 } },
      viewport,
    );
    expect(pieces).toEqual([
      expect.objectContaining({ left: 50, width: 100, borderRightWidth: 0 }),
      expect.objectContaining({ left: 150, width: 150, borderLeftWidth: 0 }),
    ]);
    vi.mocked(records[0][2].element.getBoundingClientRect).mockReturnValue(rect(350, 10, 100, 50));
    const [clipped] = measureGridSelection(
      engine,
      { sheet, selection: { x1: 2, y1: 0, x2: 2, y2: 0 } },
      viewport,
    );
    expect(clipped).toEqual(
      expect.objectContaining({
        left: 340,
        top: 24,
        width: 60,
        height: 16,
        borderRightWidth: 0,
        borderTopWidth: 0,
      }),
    );
  });

  it('scrolls the grid rather than invoking scrollIntoView on outer ancestors', () => {
    const { engine, records, viewport } = grid();
    const outerScroll = vi.fn();
    records[1][2].element.scrollIntoView = outerScroll;
    vi.mocked(records[1][2].element.getBoundingClientRect).mockReturnValue(rect(500, 500, 100, 24));
    scrollGridCellIntoView(engine, 2, 1);
    expect(viewport.scrollLeft).toBe(190);
    expect(viewport.scrollTop).toBe(104);
    expect(outerScroll).not.toHaveBeenCalled();
  });
});
