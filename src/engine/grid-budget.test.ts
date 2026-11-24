// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { type CellInput, createWorkbook } from '../model/workbook.js';
import { renderReviewMarkers } from './cell-markers.js';
import { renderCalculatedCells } from './cell-renderer.js';
import type {
  JSpreadsheetHostHandle,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';
import { windowGridRows } from './row-window.js';
import { syncSharedGridBorders } from './shared-grid-borders.js';
import { patchWorksheetData } from './worksheet-data.js';
import { syncWorksheetPresentation } from './worksheet-presentation.js';

/**
 * A 26 × 100 engine grid, the smallest the editor shows, with CE's padded data. Tests that
 * measure every cell use fewer rows: jsdom's computed style costs about a millisecond a cell.
 */
function engineGrid(data: CellInput[][], height = 100) {
  const table = document.createElement('table');
  const head = table.createTHead().insertRow();
  head.insertCell().textContent = '';
  const body = table.createTBody();
  const records: { element: HTMLTableCellElement }[][] = [];
  const rows: { element: HTMLTableRowElement }[] = [];
  const padded: CellInput[][] = [];
  for (let y = 0; y < height; y++) {
    const row = body.insertRow();
    row.insertCell().className = 'jss_row';
    rows.push({ element: row });
    records.push([]);
    padded.push([]);
    for (let x = 0; x < 26; x++) {
      const cell = row.insertCell();
      cell.dataset.x = String(x);
      cell.dataset.y = String(y);
      padded[y][x] = data[y]?.[x] ?? '';
      cell.textContent = String(padded[y][x]);
      records[y].push({ element: cell });
    }
  }
  for (let x = 0; x < 26; x++) head.insertCell().dataset.x = String(x);
  document.body.append(table);
  const engine = {
    table,
    tbody: body,
    content: table,
    rows,
    records,
    cols: [],
    headers: [],
    options: { data: padded, columns: [], rows: {}, style: {}, freezeColumns: 0 },
    getCellFromCoords: (x: number, y: number) => records[y]?.[x]?.element ?? null,
  } as unknown as JSpreadsheetWorksheetInstance;
  const handle = { getInstance: () => [engine] } as unknown as JSpreadsheetHostHandle;
  return { engine, handle, records };
}
afterEach(() => {
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

describe('SE-PERF-1.1: grid work scales with what changed', () => {
  it('repaints only cells whose value changed, and cells the engine rewrote', () => {
    const { handle, records } = engineGrid([]);
    const values = [['a', 'b', 'c']];
    renderCalculatedCells(handle, values, { raw: [['a', 'b', 'c']] });
    const nodes = records[0].map((record) => record.element.firstChild);
    renderCalculatedCells(handle, [['a', 'B', 'c']], { raw: [['a', 'B', 'c']] });
    expect(records[0][0].element.firstChild).toBe(nodes[0]);
    expect(records[0][1].element.textContent).toBe('B');
    expect(records[0][2].element.firstChild).toBe(nodes[2]);
    // The engine writes its raw value when an editor closes; the next paint restores ours.
    records[0][0].element.textContent = '=raw';
    renderCalculatedCells(handle, [['a', 'B', 'c']], { raw: [['a', 'B', 'c']] });
    expect(records[0][0].element.textContent).toBe('a');
  });

  it('skips a row whose values and input rows are the same objects as last paint', () => {
    const { handle, records } = engineGrid([]);
    const raw = [
      ['a', 'b'],
      ['c', 'd'],
    ];
    const values = [
      ['a', 'b'],
      ['c', 'd'],
    ];
    renderCalculatedCells(handle, values, { raw });
    const read = vi.spyOn(records[1][0].element.classList, 'contains');
    const child = vi.spyOn(records[1][0].element, 'firstChild', 'get');
    // Row 0 has a new values row; row 1 keeps its identity, so its cells are not visited.
    renderCalculatedCells(handle, [['A', 'b'], values[1]], { raw });
    expect(records[0][0].element.textContent).toBe('A');
    expect(read).not.toHaveBeenCalled();
    expect(child).not.toHaveBeenCalled();
    // A cell the engine rewrote is painted again even in an unchanged row.
    records[1][1].element.textContent = '=raw';
    renderCalculatedCells(handle, [['A', 'b'], values[1]], { raw });
    expect(records[1][1].element.textContent).toBe('d');
    const text = records[1][1].element.firstChild;
    if (text) text.textContent = '=native';
    renderCalculatedCells(handle, [['A', 'b'], values[1]], { raw });
    expect(records[1][1].element.textContent).toBe('d');
  });

  it('paints current input dimensions while a differently shaped result is pending', () => {
    const { handle, records } = engineGrid([]);
    const sheetId = 'sheet';
    renderCalculatedCells(handle, [['computed'], ['old row']], {
      raw: [['=A2', '=A2']],
      presentation: { sheetId, readyCells: new Set([JSON.stringify([sheetId, 0, 0])]) },
    });
    expect(records[0][0].element.textContent).toBe('computed');
    expect(records[0][1].element.textContent).toBe('…');
    expect(records[1][0].element.textContent).toBe('');
  });

  it('paints only attached rows and keeps stable review markers on detached rows', () => {
    const { engine, handle, records } = engineGrid([], 200);
    const values = Array.from({ length: 200 }, () => ['result']);
    const stop = windowGridRows(engine, () => renderCalculatedCells(handle, values));
    renderCalculatedCells(handle, values);
    expect(records[199][0].element.isConnected).toBe(false);
    expect(records[199][0].element.textContent).toBe('');
    const marks = {
      comments: [
        { id: 'comment', range: 'A200', sheetId: 'sheet', state: 'open' as const, entries: [] },
      ],
    };
    expect(renderReviewMarkers(handle, 'sheet', marks)).toBe(true);
    expect(renderReviewMarkers(handle, 'sheet', marks)).toBe(false);
    const viewport = engine.content as HTMLElement;
    viewport.scrollTop = 199 * 24;
    viewport.dispatchEvent(new Event('scroll'));
    expect(records[199][0].element.isConnected).toBe(true);
    expect(records[199][0].element.textContent).toBe('result');
    stop();
  });

  it('paints a draft row again once its draft ends', () => {
    const { handle, records } = engineGrid([]);
    const raw = [['a', 'b']];
    const values = [['a', 'b']];
    const formats = {};
    const presentation = { sheetId: 's', draft: { x: 1, y: 0, text: 'typed' } };
    renderCalculatedCells(handle, values, {
      raw,
      unavailableText: '…',
      formats,
      showFormulas: false,
      presentation,
    });
    expect(records[0][1].element.textContent).toBe('typed');
    renderCalculatedCells(handle, values, {
      raw,
      unavailableText: '…',
      formats,
      showFormulas: false,
      presentation: { sheetId: 's' },
    });
    expect(records[0][1].element.textContent).toBe('b');
  });

  it('compares calculated numbers by content, so a new result repaints nothing', () => {
    const { handle, records } = engineGrid([]);
    const number = () => ({
      kind: 'number' as const,
      numerator: '3',
      denominator: '2',
      approximate: false,
    });
    renderCalculatedCells(handle, [[number()]], { raw: [['=3/2']] });
    const node = records[0][0].element.firstChild;
    renderCalculatedCells(handle, [[number()]], { raw: [['=3/2']] });
    expect(records[0][0].element.firstChild).toBe(node);
    expect(records[0][0].element.textContent).toBe('1.5');
  });

  it('repaints every cell when display options change', () => {
    const { handle, records } = engineGrid([]);
    const third = { kind: 'number' as const, numerator: '1', denominator: '3', approximate: false };
    renderCalculatedCells(handle, [[third]], {
      display: { significantDigits: 32 },
      raw: [['=1/3']],
    });
    renderCalculatedCells(handle, [[third]], {
      display: { significantDigits: 3 },
      raw: [['=1/3']],
    });
    expect(records[0][0].element.textContent).toBe('0.333');
  });

  it('writes changed values into the engine in place and refuses large or reshaped changes', () => {
    const { engine, records } = engineGrid([['1', '2']]);
    const row = records[0][1].element;
    const sheet = createWorkbook([['1', '5']]).worksheets[0];
    const data = engine.options?.data as CellInput[][];
    expect(patchWorksheetData(engine, sheet)).toBe(true);
    expect(data[0]).toEqual(['1', '5', ...Array(24).fill('')]);
    expect(records[0][1].element).toBe(row);
    expect(row.textContent).toBe('5');
    // A grid of another extent, or more than 2,000 changed cells, rebuilds the rows instead.
    expect(patchWorksheetData(engine, createWorkbook([Array(27).fill('x')]).worksheets[0])).toBe(
      false,
    );
    const many = Array.from({ length: 100 }, () => Array<CellInput>(26).fill('x'));
    expect(patchWorksheetData(engine, createWorkbook(many).worksheets[0])).toBe(false);
    expect(data[1][0]).toBe('');
  });

  it('measures only cells with an inline border or a clipboard edge, and their neighbors', () => {
    const { engine, records } = engineGrid([]);
    records[1][1].element.style.border = '1px solid red';
    records[1][2].element.style.border = '1px solid red';
    const measured = vi.spyOn(window, 'getComputedStyle');
    syncSharedGridBorders(engine);
    expect(records[1][2].element.style.borderLeftColor).toBe('transparent');
    expect(records[1][1].element.style.borderRightColor).toBe('red');
    // Two bordered cells, their right/bottom neighbors and their left/top neighbors.
    expect(measured.mock.calls.length).toBeLessThan(12);
  });

  it('restyles only cells whose authored style changed while zoom is unchanged', () => {
    const { engine, records } = engineGrid([['1', '2']], 10);
    const sheet = createWorkbook([['1', '2']]).worksheets[0];
    syncWorksheetPresentation(
      engine,
      { previous: null, next: sheet },
      { readOnly: false, zoom: 125 },
    );
    expect(records[5][5].element.style.padding).not.toBe('');
    const measured = vi.spyOn(window, 'getComputedStyle');
    syncWorksheetPresentation(
      engine,
      { previous: sheet, next: sheet },
      { readOnly: false, zoom: 125 },
    );
    expect(measured).not.toHaveBeenCalled();
    const styled = { ...sheet, style: { B1: 'font-weight: bold' } };
    syncWorksheetPresentation(
      engine,
      { previous: sheet, next: styled },
      { readOnly: false, zoom: 125 },
    );
    expect(records[0][1].element.style.fontWeight).toBe('bold');
    expect(measured.mock.calls.length).toBe(1);
    syncWorksheetPresentation(
      engine,
      { previous: styled, next: styled },
      { readOnly: false, zoom: 100 },
    );
    expect(records[5][5].element.style.padding).toBe('');
  });
});
