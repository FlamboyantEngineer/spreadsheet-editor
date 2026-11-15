// @vitest-environment jsdom
import { useState } from 'react';
import { cleanup, render, screen, within } from '@testing-library/react';
import { UiProvider } from '@flamboyantengineer/ui-components/core';
import { userEvent } from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { MenuEntry, MenuItemEntry } from '@flamboyantengineer/ui-components/menu';
import { createWorkbook, type Workbook } from '../model/workbook.js';
import type { WorkbookCommandResult } from '../model/operations.js';
import type { SpreadsheetEditorController } from './use-editor.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { editorCommandMetadata } from './editor-commands.js';
import { editorMenus } from './editor-menus.js';
import { EditorToolbar } from './editor-toolbar.js';
import { FULL_CAPABILITIES } from '../host/capabilities.js';
import { createMemoryHost } from '../host/memory-host.js';
import { createWorkbookSession } from '../host/workbook-session.js';

function fixture(
  props: Partial<SpreadsheetEditorProps> = {},
  overrides: Partial<SpreadsheetEditorController> = {},
) {
  const workbook = createWorkbook([['1', '2']]);
  let mutationBase = workbook;
  const mutateWithDraft = vi.fn(
    (transform: (base: Workbook) => Workbook | WorkbookCommandResult) => {
      const result = transform(mutationBase);
      return 'workbook' in result ? result.workbook : result;
    },
  );
  const editor = {
    props,
    workbook,
    sheet: workbook.worksheets[0],
    mode: 'author',
    capabilities: FULL_CAPABILITIES,
    locked: false,
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

function buttonText(button: Element) {
  const copy = button.cloneNode(true) as Element;
  copy.querySelectorAll('[hidden]').forEach((description) => description.remove());
  return copy.textContent;
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
function openSession() {
  const host = createMemoryHost({ workbooks: { book: createWorkbook() } });
  return createWorkbookSession({ workbooks: host.workbooks, workbookId: 'book' });
}

describe('SE-CHROME-1 / SE-CONTEXT-1: shared command presentation', () => {
  it('gives every app-owned context/menu action and submenu a public icon, with only supported shortcuts', () => {
    const { editor } = fixture({
      session: openSession(),
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
    expect(mutateWithDraft.mock.results[0].value.worksheets[0].data).toEqual([['1E-25', '=1-A1']]);
    expect(mutateWithDraft.mock.results[0].value.worksheets[0].minDimensions).toEqual([26, 101]);
    menus.appendColumn();
    expect(mutateWithDraft.mock.results[1].value.worksheets[0].data).toEqual([['1E-25', '=1-A1']]);
    expect(mutateWithDraft.mock.results[1].value.worksheets[0].minDimensions).toEqual([27, 100]);
    action(menus.view, 'freeze').onAction!();
    const frozen = mutateWithDraft.mock.results[2].value;
    expect(frozen.worksheets[0].data[0][0]).toBe('1E-25');
    expect(frozen.worksheets[0].freezeColumns).toBe(1);
    expect(composed.worksheets[0].freezeColumns).toBeUndefined();
  });
});

describe('SE-CHROME-4: embedded icon toolbar', () => {
  it('renders new border and merge glyphs through public provider overrides and retains their commands', async () => {
    const user = userEvent.setup();
    const { editor } = fixture({ chrome: 'embedded' });
    render(
      <UiProvider
        icons={{
          'table-border-top': () => <svg data-testid="provided-border-top" aria-hidden="true" />,
          'merge-cells': () => <svg data-testid="provided-merge" aria-hidden="true" />,
        }}
      >
        <EditorToolbar editor={editor} />
      </UiProvider>,
    );
    const merge = screen.getByRole('button', { name: 'Merge cells' });
    expect(merge.querySelector('[data-testid="provided-merge"]')).toBeTruthy();
    await user.click(merge);
    expect(editor.merge).toHaveBeenCalledWith();
    await user.click(screen.getByRole('button', { name: 'Borders' }));
    const top = screen.getByRole('menuitem', { name: 'Top borders' });
    expect(top.querySelector('[data-testid="provided-border-top"]')).toBeTruthy();
    const outside = screen.getByRole('menuitem', { name: 'Outside borders' });
    expect(outside.querySelectorAll('svg path').length).toBeGreaterThan(1);
    expect(outside.querySelector('path[stroke-dasharray]')).toBeTruthy();
    await user.click(top);
    expect(editor.border).toHaveBeenCalledWith('top');
  });

  it('renders named glyph-only controls using the public toolbar layout, including unit and Save actions', () => {
    const { editor } = fixture(
      { chrome: 'embedded', session: openSession() },
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
      'Text color',
      'Cell fill',
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
      expect(buttonText(button), name).toBe('');
    }
    for (const button of container.querySelectorAll('button')) expect(buttonText(button)).toBe('');
    expect(screen.queryByRole('menubar')).toBeNull();
    expect(screen.queryByRole('combobox')).toBeNull();
  });

  it('keeps save retry available through the icon action after a failed embedded save', async () => {
    const user = userEvent.setup();
    const { editor } = fixture(
      { chrome: 'embedded', session: openSession() },
      { savePhase: 'error' },
    );
    const { container } = render(<EditorToolbar editor={editor} />);
    for (const button of container.querySelectorAll('button')) expect(buttonText(button)).toBe('');
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

describe('SE-CHROME-7: public text formatting controls', () => {
  it('steps a font size between presets with the public glyphs and accepts a custom fractional size', async () => {
    const user = userEvent.setup();
    const { editor } = fixture({ chrome: 'full' }, { style: { 'font-size': '13px' } });
    render(
      <UiProvider
        icons={{
          'increase-font-size': () => <svg data-testid="provided-size-up" aria-hidden="true" />,
          'decrease-font-size': () => <svg data-testid="provided-size-down" aria-hidden="true" />,
        }}
      >
        <EditorToolbar editor={editor} />
      </UiProvider>,
    );
    const input = screen.getByRole('textbox', { name: 'Font size' });
    const increase = screen.getByRole('button', { name: 'Increase Font size' });
    const decrease = screen.getByRole('button', { name: 'Decrease Font size' });
    expect(increase.querySelector('[data-testid="provided-size-up"]')).toBeTruthy();
    expect(decrease.querySelector('[data-testid="provided-size-down"]')).toBeTruthy();
    await user.click(increase);
    expect(editor.format).toHaveBeenLastCalledWith({ 'font-size': '14px' });
    await user.clear(input);
    await user.type(input, '13.5{Enter}');
    expect(editor.format).toHaveBeenLastCalledWith({ 'font-size': '13.5px' });
    expect(screen.queryByRole('button', { name: /restore|apply|cancel/i })).toBeNull();
  });

  it('opens the public preset field from the embedded font icon and closes it when editing is disabled', async () => {
    const user = userEvent.setup();
    const { editor } = fixture({ chrome: 'embedded' });
    const { rerender } = render(<EditorToolbar editor={editor} />);
    await user.click(screen.getByRole('button', { name: 'Font size' }));
    const dialog = screen.getByRole('dialog', { name: 'Font size' });
    await user.click(within(dialog).getByRole('textbox', { name: 'Font size' }));
    await user.keyboard('{ArrowUp}');
    expect(editor.format).toHaveBeenLastCalledWith({ 'font-size': '14px' });
    rerender(<EditorToolbar editor={{ ...editor, canWrite: false }} />);
    expect(screen.queryByRole('dialog', { name: 'Font size' })).toBeNull();
    expect(screen.getByRole('button', { name: 'Font size' }).hasAttribute('disabled')).toBe(true);
  });

  for (const chrome of ['full', 'embedded'] as const)
    it(`shows the public color bars and applies preset, custom and cleared colors in ${chrome} mode`, async () => {
      const user = userEvent.setup();
      const { editor } = fixture(
        { chrome },
        { style: { color: '#175cd3', 'background-color': '#fef0c7' } },
      );
      function ControlledToolbar() {
        const [style, setStyle] = useState(editor.style);
        return (
          <EditorToolbar
            editor={{
              ...editor,
              style,
              format: (patch) => {
                editor.format(patch);
                setStyle(
                  (current) =>
                    Object.fromEntries(
                      Object.entries({ ...current, ...patch }).filter(
                        ([, value]) => value !== null,
                      ),
                    ) as Record<string, string>,
                );
              },
            }}
          />
        );
      }
      render(
        <UiProvider
          icons={{
            'font-color': () => <svg data-testid="provided-text-color" aria-hidden="true" />,
            'background-color': () => <svg data-testid="provided-fill-color" aria-hidden="true" />,
          }}
        >
          <ControlledToolbar />
        </UiProvider>,
      );
      const textColor = screen.getByRole('button', { name: 'Text color' });
      const fillColor = screen.getByRole('button', { name: 'Cell fill' });
      expect(textColor.querySelector('[data-testid="provided-text-color"]')).toBeTruthy();
      expect(fillColor.querySelector('[data-testid="provided-fill-color"]')).toBeTruthy();
      expect(textColor.querySelector('.ui-color-swatch-button__swatch')).toBeTruthy();
      await user.click(textColor);
      await user.click(screen.getByRole('option', { name: '#b42318' }));
      expect(editor.format).toHaveBeenLastCalledWith({ color: '#b42318' });
      expect(screen.queryByRole('dialog', { name: 'Text color' })).toBeNull();
      await user.click(fillColor);
      const input = screen.getByRole('textbox', { name: 'Custom cell fill' });
      await user.clear(input);
      await user.type(input, '#336699{Enter}');
      expect(editor.format).toHaveBeenLastCalledWith({ 'background-color': '#336699' });
      await user.click(screen.getByRole('button', { name: 'Clear cell fill' }));
      expect(editor.format).toHaveBeenLastCalledWith({ 'background-color': null });
    });

  it('disables colors and closes an open palette when write permission is removed', async () => {
    const user = userEvent.setup();
    const { editor } = fixture({ chrome: 'embedded' });
    const { rerender } = render(<EditorToolbar editor={editor} />);
    await user.click(screen.getByRole('button', { name: 'Text color' }));
    expect(screen.getByRole('dialog', { name: 'Text color' })).toBeTruthy();
    rerender(<EditorToolbar editor={{ ...editor, canWrite: false }} />);
    expect(screen.queryByRole('dialog', { name: 'Text color' })).toBeNull();
    expect(screen.getByRole('button', { name: 'Text color' }).hasAttribute('disabled')).toBe(true);
    expect(screen.getByRole('button', { name: 'Cell fill' }).hasAttribute('disabled')).toBe(true);
    expect(editor.format).not.toHaveBeenCalled();
  });
});
