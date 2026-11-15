// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { clearCalculatedPresentation, renderCalculatedCells } from './cell-renderer.js';
import type { JSpreadsheetHostHandle } from './jspreadsheet-types.js';

function grid() {
  const cells = [document.createElement('td'), document.createElement('td')];
  cells.forEach((cell) => document.body.append(cell));
  const worksheet = { getCellFromCoords: (x: number) => cells[x] };
  const handle = { getInstance: () => [worksheet] } as unknown as JSpreadsheetHostHandle;
  return { cells, handle };
}
const value = { kind: 'number' as const, numerator: '1', denominator: '1', approximate: false };
const key = JSON.stringify(['sheet', 0, 0]);
const publication = { inputKey: 'accepted', revision: 2 };
afterEach(() => document.body.replaceChildren());

describe('SE-HOST-1.1: current presentation and protected draft painting', () => {
  it('paints masked nonready results as pending rather than an empty calculated value', () => {
    const { handle, cells } = grid();
    renderCalculatedCells(handle, [[null, 'unrelated']], {
      raw: [['=A2', '=1']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: {
        sheetId: 'sheet',
        readyCells: new Set([JSON.stringify(['sheet', 1, 0])]),
      },
    });
    expect(cells[0].textContent).toBe('…');
    expect(cells[1].textContent).toBe('unrelated');
  });

  it('paints a protected draft without changing raw inputs or unrelated calculated cells', () => {
    const { handle, cells } = grid();
    const raw = [['=1/3', '=1+1']];
    renderCalculatedCells(handle, [[value, 'unrelated']], {
      raw,
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: {
        sheetId: 'sheet',
        draft: { x: 0, y: 0, text: '=1/7' },
      },
    });
    expect(cells[0].textContent).toBe('=1/7');
    expect(cells[1].textContent).toBe('unrelated');
    expect(raw).toEqual([['=1/3', '=1+1']]);
    renderCalculatedCells(handle, [[value, 'unrelated']], { raw });
    expect(cells[0].textContent).toBe('1');
  });

  it('accepts matching presentation and clears old titles/decorations on a stale publication', () => {
    const { handle, cells } = grid();
    const hint = {
      text: 'Captured value',
      tooltip: 'Source pin',
      ariaDescription: 'Pinned',
      decoration: 'external',
    };
    const options = {
      sheetId: 'sheet',
      publication,
      cellPresentation: { ...publication, cells: { [key]: hint } },
    };
    renderCalculatedCells(handle, [[value]], {
      raw: [['=FOO()']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: options,
    });
    expect(cells[0].textContent).toBe('Captured value');
    expect(cells[0].title).toBe('Source pin');
    expect(cells[0].getAttribute('aria-description')).toBe('Pinned');
    expect(cells[0].dataset.seDecoration).toBe('external');
    renderCalculatedCells(handle, [[value]], {
      raw: [['=FOO()']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: {
        ...options,
        publication: { inputKey: 'changed', revision: 3 },
      },
    });
    expect(cells[0].textContent).toBe('1');
    expect(cells[0].hasAttribute('title')).toBe(false);
    expect(cells[0].hasAttribute('aria-description')).toBe(false);
    expect(cells[0].dataset.seDecoration).toBeUndefined();
  });

  it('creates stable engine-owned targets for current React content and releases them on pending/draft', () => {
    const { handle, cells } = grid();
    const renderCellContent = vi.fn(() => 'Rich value');
    const options = { sheetId: 'sheet', publication, renderCellContent };
    const first = renderCalculatedCells(handle, [[value]], {
      raw: [['=FOO()']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: options,
    });
    expect(first[0].target.parentElement).toBe(cells[0]);
    expect(first[0].target.className).toBe('se-editor__cell-content');
    expect(renderCellContent).toHaveBeenLastCalledWith(
      expect.objectContaining({
        cellKey: key,
        rawInput: '=FOO()',
        value,
        inputKey: 'accepted',
        revision: 2,
        calculatedText: '1',
      }),
    );
    const next = renderCalculatedCells(handle, [[value]], {
      raw: [['=FOO()']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: options,
    });
    expect(next[0].target).toBe(first[0].target);
    renderCellContent.mockClear();
    const pending = renderCalculatedCells(handle, [['…']], {
      raw: [['=FOO()']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: {
        ...options,
        readyCells: new Set(),
      },
    });
    expect(pending).toEqual([]);
    expect(renderCellContent).not.toHaveBeenCalled();
    expect(cells[0].textContent).toBe('…');
    renderCalculatedCells(handle, [[value]], {
      raw: [['=FOO()']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: options,
    });
    clearCalculatedPresentation(handle);
    expect(cells[0].querySelector('.se-editor__cell-content')).toBeNull();
  });

  it('uses undefined for default content and null for intentionally empty React content', () => {
    const { handle, cells } = grid();
    const defaultPortals = renderCalculatedCells(handle, [[value]], {
      raw: [['1']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: {
        sheetId: 'sheet',
        publication,
        renderCellContent: () => undefined,
      },
    });
    expect(defaultPortals).toEqual([]);
    expect(cells[0].textContent).toBe('1');
    const emptyPortals = renderCalculatedCells(handle, [[value]], {
      raw: [['1']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: {
        sheetId: 'sheet',
        publication,
        renderCellContent: () => null,
      },
    });
    expect(emptyPortals).toHaveLength(1);
    expect(emptyPortals[0].content).toBeNull();
    expect(cells[0].textContent).toBe('');
  });

  it('uses a suggested unit only when there is no explicit authored format', () => {
    const { handle, cells } = grid();
    const length = { ...value, dimension: [1, 0, 0, 0, 0, 0, 0] as const };
    const options = {
      sheetId: 'sheet',
      publication,
      cellPresentation: { ...publication, cells: { [key]: { defaultUnit: 'cm' } } },
    };
    renderCalculatedCells(handle, [[length]], {
      raw: [['1 m']],
      unavailableText: '…',
      formats: {},
      showFormulas: false,
      presentation: options,
    });
    expect(cells[0].textContent).toBe('100 cm');
    renderCalculatedCells(handle, [[length]], {
      raw: [['1 m']],
      unavailableText: '…',
      formats: { A1: { kind: 'unit', unit: 'mm' } },
      showFormulas: false,
      presentation: options,
    });
    expect(cells[0].textContent).toBe('1000 mm');
    renderCalculatedCells(handle, [[value]], {
      raw: [['1']],
      unavailableText: '…',
      formats: { A1: { kind: 'number', decimals: 2 } },
      showFormulas: false,
      presentation: options,
    });
    expect(cells[0].textContent).toBe('1.00');
  });
});
