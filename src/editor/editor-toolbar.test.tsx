// @vitest-environment jsdom
import { cleanup, render, screen, within } from '@testing-library/react';
import { userEvent } from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { MenuEntry, MenuItemEntry } from '@flamboyantengineer/ui-components/menu';
import { createWorkbook, type Workbook } from '../model/workbook.js';
import type { SpreadsheetEditorController } from './use-editor.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { editorCommandMetadata } from './editor-commands.js';
import { editorMenus } from './editor-menus.js';
import { EditorToolbar } from './editor-toolbar.js';

function fixture(
  props: Partial<SpreadsheetEditorProps> = {},
  overrides: Partial<SpreadsheetEditorController> = {},
) {
  const workbook = createWorkbook([['1', '2']]);
  let mutationBase = workbook;
  const mutateWithDraft = vi.fn((transform: (base: Workbook) => Workbook) =>
    transform(mutationBase),
  );
  const editor = {
    props,
    workbook,
    sheet: workbook.worksheets[0],
    mode: 'author',
    canWrite: true,
    canEdit: true,
    canComment: true,
    style: {},
    cellFormat: { kind: 'general' },
    selection: { x1: 0, x2: 1, y1: 0, y2: 0 },
    history: { past: [], future: [] },
    calculation: { pending: false },
    peers: [],
    grid: { current: null },
    savePhase: 'dirty',
    draftDirty: false,
    attempt: vi.fn((action: () => unknown) => action()),
    format: vi.fn(),
    toggleFormat: vi.fn(),
    undo: vi.fn(),
    redo: vi.fn(),
    copy: vi.fn(),
    paste: vi.fn(),
    clear: vi.fn(),
    structure: vi.fn(),
    border: vi.fn(),
    merge: vi.fn(),
    sheetCommand: vi.fn(),
    beginComment: vi.fn(),
    setShowFind: vi.fn(),
    setShowContributors: vi.fn(),
    setShowGridlines: vi.fn(),
    setShowFormulas: vi.fn(),
    openPanel: vi.fn(),
    openDialog: vi.fn(),
    setMode: vi.fn(),
    save: vi.fn(),
    refresh: vi.fn(),
    importExcel: vi.fn(),
    exportExcel: vi.fn(),
    insert: vi.fn(),
    mutate: vi.fn(),
    mutateWithDraft,
    ...overrides,
  } as unknown as SpreadsheetEditorController;
  return {
    editor,
    mutateWithDraft,
    setMutationBase: (base: Workbook) => {
      mutationBase = base;
    },
  };
}

function entries(items: readonly MenuEntry[]): MenuItemEntry[] {
  return items.flatMap((item) =>
    item.type === 'separator'
      ? []
      : item.type === 'section'
        ? entries(item.items)
        : [item, ...entries(item.items ?? [])],
  );
}

function action(items: MenuEntry[], id: string) {
  return entries(items).find((item) => item.id === id)!;
}

beforeEach(() => {
  // JSDOM omits CSS.escape, which React Aria uses to locate menu keys.
  vi.stubGlobal('CSS', {
    ...globalThis.CSS,
    escape: (value: string) =>
      Array.from(value, (character) => `\\${character.codePointAt(0)!.toString(16)} `).join(''),
  });
});
afterEach(() => {
  cleanup();
  vi.unstubAllGlobals();
});
describe('SE-CHROME-1 / SE-CONTEXT-1: shared command presentation', () => {
  it('gives every app-owned context/menu action and submenu a public icon, with only supported shortcuts', () => {
    const { editor } = fixture({
      onSave: vi.fn(),
      onRefresh: vi.fn(),
      onImportXlsx: vi.fn(),
      onExportXlsx: vi.fn(),
      onInsert: vi.fn(),
      panels: [{ id: 'workflow', label: 'Workflow', content: null }],
    });
    const menus = editorMenus(editor);
    const all = Object.values(menus).flatMap((value) =>
      Array.isArray(value) ? entries(value) : [],
    );
    for (const item of all) {
      expect(item.icon, String(item.id)).toBeTruthy();
      if (item.id !== 'pane-workflow')
        expect(item.icon, String(item.id)).not.toBe(editorCommandMetadata('pane-workflow').icon);
      if (item.shortcut)
        expect([
          'Ctrl+Z',
          'Ctrl+Y',
          'Ctrl+X',
          'Ctrl+C',
          'Ctrl+V',
          'Ctrl+F',
          'Ctrl+S',
          'Ctrl+B',
          'Ctrl+I',
          'Ctrl+Alt+M',
          'Delete',
        ]).toContain(item.shortcut);
    }
    expect(action(menus.context, 'copy').shortcut).toBe('Ctrl+C');
    expect(action(menus.context, 'context-comment').shortcut).toBe('Ctrl+Alt+M');
    expect(action(menus.review, 'new-comment').shortcut).toBe('Ctrl+Alt+M');
    expect(action(menus.rowContext, 'row-delete').shortcut).toBe('Delete');
    expect(action(menus.columnContext, 'col-delete').shortcut).toBe('Delete');
    expect(action(menus.context, 'delete-rows').shortcut).toBeUndefined();
    expect(action(menus.format, 'strike').shortcut).toBeUndefined();
    expect(menus.align).toHaveLength(9);
    expect(menus.borders).toHaveLength(7);
  });

  it('keeps review/reading structural and mutation commands disabled while copying and menus remain available', () => {
    const { editor } = fixture(
      { features: { search: false } },
      { canWrite: false, canEdit: false, canComment: false },
    );
    const menus = editorMenus(editor);
    for (const id of [
      'cut',
      'paste',
      'clear',
      'row-above',
      'delete-rows',
      'bold',
      'merge',
      'context-comment',
    ])
      expect(action(menus.context, id).isDisabled, id).toBe(true);
    expect(action(menus.context, 'copy').isDisabled).toBe(false);
    expect(action(menus.edit, 'find').isDisabled).toBe(true);
    expect(action(menus.view, 'pane-history').isDisabled).toBe(false);
  });

  it('applies structural toolbar actions and freezing to the draft-composed workbook base', () => {
    const { editor, mutateWithDraft, setMutationBase } = fixture();
    const composed = createWorkbook([['1E-25', '=1-A1']]);
    setMutationBase(composed);
    const menus = editorMenus(editor);
    menus.appendRow();
    expect(mutateWithDraft.mock.results[0].value.worksheets[0].data).toEqual([
      ['1E-25', '=1-A1'],
      [null, null],
    ]);
    menus.appendColumn();
    expect(mutateWithDraft.mock.results[1].value.worksheets[0].data).toEqual([
      ['1E-25', '=1-A1', null],
    ]);
    action(menus.view, 'freeze').onAction!();
    const frozen = mutateWithDraft.mock.results[2].value;
    expect(frozen.worksheets[0].data[0][0]).toBe('1E-25');
    expect(frozen.worksheets[0].freezeColumns).toBe(1);
    expect(composed.worksheets[0].freezeColumns).toBeUndefined();
  });
});

describe('SE-CHROME-4: embedded icon toolbar', () => {
  it('renders named glyph-only controls using the public toolbar layout, including unit and Save actions', () => {
    const { editor } = fixture(
      { chrome: 'embedded', onSave: vi.fn() },
      { cellFormat: { kind: 'unit', unit: 'mm' } },
    );
    const { container } = render(<EditorToolbar editor={editor} />);
    expect(screen.getByRole('toolbar', { name: 'Spreadsheet formatting' }).className).toContain(
      'ui-toolbar-row',
    );
    for (const name of [
      'Spreadsheet actions',
      'Editing mode',
      'Save',
      'Font family',
      'Font size',
      'Number format',
      'Display unit',
      'Increase decimal places',
      'Decrease decimal places',
      'Align cells',
      'Borders',
      'Wrap text',
      'Merge cells',
      'Unmerge cells',
    ]) {
      const button = screen.getByRole('button', { name });
      expect(button.querySelector('svg'), name).toBeTruthy();
      expect(button.textContent, name).toBe('');
    }
    for (const button of container.querySelectorAll('button')) expect(button.textContent).toBe('');
    expect(screen.queryByRole('menubar')).toBeNull();
    expect(screen.queryByRole('combobox')).toBeNull();
  });

  it('keeps save retry available through the icon action after a failed embedded save', async () => {
    const user = userEvent.setup();
    const { editor } = fixture({ chrome: 'embedded', onSave: vi.fn() }, { savePhase: 'error' });
    const { container } = render(<EditorToolbar editor={editor} />);
    for (const button of container.querySelectorAll('button')) expect(button.textContent).toBe('');
    await user.click(screen.getByRole('button', { name: 'Save' }));
    expect(editor.save).toHaveBeenCalledOnce();
  });

  it('opens labelled font and format choices and preserves the selected value and action', async () => {
    const user = userEvent.setup();
    const { editor } = fixture({ chrome: 'embedded' });
    render(<EditorToolbar editor={editor} />);
    await user.click(screen.getByRole('button', { name: 'Font family' }));
    expect(
      screen.getByRole('menuitemcheckbox', { name: 'sans-serif' }).getAttribute('aria-checked'),
    ).toBe('true');
    await user.click(screen.getByRole('menuitemcheckbox', { name: 'Georgia' }));
    expect(editor.format).toHaveBeenCalledWith({ 'font-family': 'Georgia' });
    await user.click(screen.getByRole('button', { name: 'Number format' }));
    await user.click(screen.getByRole('menuitemcheckbox', { name: 'Scientific' }));
    expect(editor.format).toHaveBeenCalledWith({}, { kind: 'scientific' });
  });

  it('keeps full editor actions in one icon toolbar and runs the moved workbook commands', async () => {
    const user = userEvent.setup();
    const { editor, mutateWithDraft } = fixture({ chrome: 'full', onRefresh: vi.fn() });
    render(<EditorToolbar editor={editor} />);
    const toolbar = within(screen.getByRole('toolbar', { name: 'Spreadsheet formatting' }));
    expect(screen.getAllByRole('toolbar')).toHaveLength(1);
    for (const name of [
      'Add row',
      'Add column',
      'Find',
      'Wrap text',
      'Merge cells',
      'Unmerge cells',
      'Align cells',
      'Borders',
      'Refresh data',
    ]) {
      const button = toolbar.getByRole('button', { name });
      expect(button.querySelector('svg'), name).toBeTruthy();
      expect(button.textContent, name).toBe('');
    }
    await user.click(toolbar.getByRole('button', { name: 'Add row' }));
    expect(mutateWithDraft).toHaveBeenCalledOnce();
    await user.click(toolbar.getByRole('button', { name: 'Find' }));
    expect(editor.setShowFind).toHaveBeenCalledWith(true);
    await user.click(toolbar.getByRole('button', { name: 'Comments (0)' }));
    expect(editor.openPanel).toHaveBeenCalledWith('comments');
  });

  it('keeps the full labelled menubar and field controls free of restore/apply/cancel actions after a value changes', () => {
    const { editor } = fixture({ chrome: 'full' });
    const { rerender } = render(<EditorToolbar editor={editor} />);
    rerender(
      <EditorToolbar
        editor={{
          ...editor,
          style: { 'font-family': 'Georgia', 'font-size': '18px' },
          cellFormat: { kind: 'number', decimals: 5 },
        }}
      />,
    );
    expect(screen.getByRole('menubar', { name: 'Spreadsheet menu' })).toBeTruthy();
    expect(screen.getByRole('button', { name: /Font family/ })).toBeTruthy();
    expect(screen.queryByRole('button', { name: /restore|apply|cancel/i })).toBeNull();
  });
});
