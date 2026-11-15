// @vitest-environment jsdom
import { waitFor } from '@testing-library/react';
import { afterEach, describe, expect, it } from 'vitest';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';
import {
  observeGridBorderCopies,
  restoreSharedGridBorders,
  syncSharedGridBorders,
} from './shared-grid-borders.js';

function fixture(merged?: 'left' | 'right') {
  const table = document.createElement('table');
  table.innerHTML =
    '<tbody><tr><td data-x="0" data-y="0">A1</td><td data-x="1" data-y="0">B1</td></tr><tr><td data-x="0" data-y="1">A2</td><td data-x="1" data-y="1">B2</td></tr></tbody>';
  const cells = Array.from(table.rows).map((row) => Array.from(row.cells));
  cells.flat().forEach((cell) => {
    cell.style.border = '1px solid red';
  });
  if (merged) {
    const x = merged === 'left' ? 0 : 1;
    cells[0][x].rowSpan = 2;
    cells[1][x].style.display = 'none';
  }
  document.body.append(table);
  const engine = {
    table,
    getCellFromCoords: (x: number, y: number) =>
      merged && x === (merged === 'left' ? 0 : 1) ? cells[0][x] : cells[y]?.[x],
  } as unknown as JSpreadsheetWorksheetInstance;
  return { engine, cells };
}
afterEach(() => {
  document.body.innerHTML = '';
});

describe('SE-UI-1: one visible pixel at authored shared borders', () => {
  it('deduplicates interior edges while retaining the complete outer perimeter', () => {
    const { engine, cells } = fixture();
    syncSharedGridBorders(engine);
    expect(cells[0][0].style.borderLeftColor).toBe('red');
    expect(cells[0][0].style.borderTopColor).toBe('red');
    expect(cells[0][0].style.borderRightColor).toBe('red');
    expect(cells[0][1].style.borderLeftColor).toBe('transparent');
    expect(cells[1][0].style.borderTopColor).toBe('transparent');
    expect(cells[1][1].style.borderRightColor).toBe('red');
    expect(cells[1][1].style.borderBottomColor).toBe('red');
    restoreSharedGridBorders(engine);
    expect(cells[0][1].style.borderLeftColor).toBe('red');
    expect(cells[1][0].style.borderTopColor).toBe('red');
  });

  it.each(['left', 'right'] as const)('preserves a continuous merged %s edge', (merged) => {
    const { engine, cells } = fixture(merged);
    syncSharedGridBorders(engine);
    if (merged === 'left') {
      expect(cells[0][0].style.borderRightColor).toBe('red');
      expect(cells[0][1].style.borderLeftColor).toBe('transparent');
      expect(cells[1][1].style.borderLeftColor).toBe('transparent');
    } else {
      expect(cells[0][1].style.borderLeftColor).toBe('red');
      expect(cells[0][0].style.borderRightColor).toBe('transparent');
      expect(cells[1][0].style.borderRightColor).toBe('transparent');
    }
  });

  it('gives a native clipboard perimeter its edge and restores dedup on dismissal', async () => {
    const { engine, cells } = fixture();
    syncSharedGridBorders(engine);
    const stop = observeGridBorderCopies(engine);
    cells[0][1].classList.add('copying-left');
    await waitFor(() => expect(cells[0][1].style.borderLeftColor).toBe('red'));
    expect(cells[0][0].style.borderRightColor).toBe('transparent');
    cells[0][1].classList.remove('copying-left');
    await waitFor(() => expect(cells[0][1].style.borderLeftColor).toBe('transparent'));
    expect(cells[0][0].style.borderRightColor).toBe('red');
    stop();
  });
});
