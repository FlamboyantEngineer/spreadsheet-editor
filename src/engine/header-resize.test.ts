// @vitest-environment jsdom
import { act, cleanup, renderHook } from '@testing-library/react';
import type { PointerEvent } from 'react';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import { useHeaderResize } from './header-resize.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';

function setup(zoom: number, axis: 'column' | 'row' = 'column') {
  const sheet = createWorkbook([['1']]).worksheets[0];
  sheet.columns = [{ width: 150.5 }];
  sheet.rows = { '0': { height: 30.5 } };
  const root = document.createElement('div');
  root.tabIndex = -1;
  const table = document.createElement('table');
  table.innerHTML =
    '<colgroup><col></colgroup><thead><tr><td data-x="0">A</td></tr></thead><tbody><tr><td data-y="0">1</td></tr></tbody>';
  root.append(table);
  document.body.append(root);
  const header =
    axis === 'column'
      ? requireDefined(table.tHead).rows[0].cells[0]
      : table.tBodies[0].rows[0].cells[0];
  vi.spyOn(header, 'getBoundingClientRect').mockReturnValue({
    left: 0,
    top: 0,
    right: 200,
    bottom: 40,
    width: 200,
    height: 40,
  } as DOMRect);
  root.setPointerCapture = vi.fn();
  root.releasePointerCapture = vi.fn();
  root.hasPointerCapture = vi.fn(() => true);
  const col = requireDefined(table.querySelector('col'));
  col.style.width = `${(150.5 * zoom) / 100}px`;
  col.setAttribute('width', String((150.5 * zoom) / 100));
  const row = table.tBodies[0].rows[0];
  row.style.height = `${(30.5 * zoom) / 100}px`;
  row.setAttribute('height', String((30.5 * zoom) / 100));
  const engine = {
    cols: [{ colElement: col }],
    rows: [{ element: row }],
  } as unknown as JSpreadsheetWorksheetInstance;
  const onChange = vi.fn(),
    onGeometryChange = vi.fn();
  const options = {
    root: { current: root },
    worksheet: sheet,
    getWorksheet: () => engine,
    canResize: true,
    zoom,
    onChange,
    onGeometryChange,
  };
  const hook = renderHook((props) => useHeaderResize(props), { initialProps: options });
  const event = (distance = 0) =>
    ({
      target: header,
      button: 0,
      pointerId: 7,
      clientX: 199 + (axis === 'column' ? distance : 0),
      clientY: 39 + (axis === 'row' ? distance : 0),
      preventDefault: vi.fn(),
      stopPropagation: vi.fn(),
    }) as unknown as PointerEvent<HTMLDivElement>;
  return { ...hook, options, event, col, row, onChange, onGeometryChange, sheet };
}
afterEach(() => {
  cleanup();
  document.body.innerHTML = '';
  vi.restoreAllMocks();
});

describe('SE-CR-2.18: resize commits logical geometry once at grid zoom', () => {
  it.each([50, 75, 100, 125, 150])('compensates a fractional column resize at %s%%', (zoom) => {
    const { result, event, col, onChange, sheet } = setup(zoom);
    act(() => {
      expect(result.current.start(event())).toBe(true);
    });
    act(() => {
      result.current.move(event(37.5));
    });
    expect(Number.parseFloat(col.style.width)).toBeCloseTo((150.5 * zoom) / 100 + 37.5);
    expect(onChange).not.toHaveBeenCalled();
    act(() => {
      result.current.finish(event(37.5));
    });
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(onChange.mock.calls[0][0].columns[0].width).toBeCloseTo(150.5 + 37.5 / (zoom / 100));
    expect(sheet.columns?.[0].width).toBe(150.5);
    expect(col.style.width).toBe(`${(150.5 * zoom) / 100}px`);
  });

  it('scales row previews and restores exact presentation on cancellation', () => {
    const { result, event, row, onChange } = setup(125, 'row');
    const original = { css: row.style.height, attribute: row.getAttribute('height') };
    act(() => {
      result.current.start(event());
    });
    act(() => {
      result.current.move(event(31.25));
    });
    expect(row.style.height).toBe('69.375px');
    act(() => {
      result.current.cancel();
    });
    expect(row.style.height).toBe(original.css);
    expect(row.getAttribute('height')).toBe(original.attribute);
    expect(onChange).not.toHaveBeenCalled();
  });

  it('cancels a live preview when zoom or edit permission changes', () => {
    const { result, rerender, options, event, col, onChange } = setup(125);
    act(() => {
      result.current.start(event());
    });
    act(() => {
      result.current.move(event(20));
    });
    rerender({ ...options, zoom: 150 });
    expect(result.current.active).toBe(false);
    expect(col.style.width).toBe('188.125px');
    act(() => {
      result.current.finish(event(20));
    });
    expect(onChange).not.toHaveBeenCalled();
    act(() => {
      result.current.start(event());
    });
    act(() => {
      result.current.move(event(20));
    });
    rerender({ ...options, zoom: 150, canResize: false });
    expect(result.current.active).toBe(false);
    expect(onChange).not.toHaveBeenCalled();
  });
});
