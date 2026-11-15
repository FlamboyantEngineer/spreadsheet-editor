// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import {
  GRID_SPACER_CLASS,
  gridRowWindow,
  revealGridRows,
  windowGridRows,
  withLazyRebuild,
} from './row-window.js';

/** An engine of `count` rows with every row attached, as it builds without lazy loading. */
function engineRows(count: number, mergeCells: Record<string, [number, number]> = {}) {
  const content = document.createElement('div');
  const table = document.createElement('table');
  table.createTHead().insertRow();
  const tbody = table.createTBody();
  content.append(table);
  document.body.append(content);
  const build = () =>
    Array.from({ length: count }, (_, y) => {
      const element = tbody.insertRow();
      element.dataset.y = String(y);
      element.style.height = '24px';
      element.insertCell().textContent = String(y + 1);
      return { element };
    });
  const engine = {
    content,
    table,
    tbody,
    rows: build(),
    cols: [{}, {}],
    options: { mergeCells, defaultRowHeight: 24 },
  } as unknown as JSpreadsheetWorksheetInstance;
  return {
    engine,
    rebuild() {
      tbody.replaceChildren();
      (engine as { rows: unknown }).rows = build();
    },
  };
}

/** The attached rows by index, and the spacer heights above and below them. */
function layout(engine: JSpreadsheetWorksheetInstance) {
  const children = [...requireDefined(engine.tbody).children] as HTMLElement[];
  const spacer = (row: HTMLElement | undefined) =>
    row?.classList.contains(GRID_SPACER_CLASS) ? (row.firstElementChild as HTMLElement) : null;
  const top = spacer(children[0]),
    bottom = spacer(children.at(-1));
  const rows = children.filter((row) => !row.classList.contains(GRID_SPACER_CLASS));
  return {
    first: Number(rows[0]?.dataset.y),
    last: Number(rows.at(-1)?.dataset.y),
    count: rows.length,
    above: top ? Number.parseFloat(top.style.height) : 0,
    below: bottom ? Number.parseFloat(bottom.style.height) : 0,
  };
}

afterEach(() => {
  document.body.replaceChildren();
});

describe('SE-PERF-1.1: the grid attaches only the rows near the viewport', () => {
  it('keeps the detached rows as spacer height above and below the window', () => {
    const { engine } = engineRows(300);
    windowGridRows(engine, () => {});
    expect(layout(engine)).toEqual({ first: 0, last: 19, count: 20, above: 0, below: 280 * 24 });
    expect(gridRowWindow(engine)).toEqual({ start: 0, end: 20 });
  });

  it('attaches a far row around itself and reports the rows it attached', () => {
    const { engine } = engineRows(300);
    const attached = vi.fn();
    windowGridRows(engine, attached);
    attached.mockClear();
    revealGridRows(engine, 250);
    expect(layout(engine)).toEqual({
      first: 230,
      last: 270,
      count: 41,
      above: 230 * 24,
      below: 29 * 24,
    });
    expect(attached).toHaveBeenCalledWith(Array.from({ length: 41 }, (_, index) => 230 + index));
  });

  it('attaches every row of a merge that crosses the window edge', () => {
    // A10 spans 30 rows, from row 9 to row 38.
    const { engine } = engineRows(300, { A10: [1, 30] });
    windowGridRows(engine, () => {});
    expect(layout(engine)).toMatchObject({ first: 0, last: 38, below: 261 * 24 });
  });

  it('attaches every row of a small worksheet', () => {
    const { engine } = engineRows(150);
    windowGridRows(engine, () => {});
    expect(layout(engine)).toEqual({ first: 0, last: 149, count: 150, above: 0, below: 0 });
    expect(gridRowWindow(engine)).toEqual({ start: 0, end: 150 });
  });

  it('takes over the rows of a rebuild and builds them lazily', () => {
    const { engine, rebuild } = engineRows(300);
    windowGridRows(engine, () => {});
    revealGridRows(engine, 250);
    withLazyRebuild(engine, () => {
      expect((engine.options as { lazyLoading?: boolean }).lazyLoading).toBe(true);
      rebuild();
    });
    expect((engine.options as { lazyLoading?: boolean }).lazyLoading).toBe(false);
    expect(layout(engine)).toEqual({ first: 0, last: 19, count: 20, above: 0, below: 280 * 24 });
  });

  it('gives hidden rows no height', () => {
    const { engine } = engineRows(300);
    (requireDefined(engine.rows)[299].element as HTMLElement).style.display = 'none';
    windowGridRows(engine, () => {});
    expect(layout(engine).below).toBe(279 * 24);
  });
});
