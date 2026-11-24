// @vitest-environment jsdom
import { UiProvider } from '@flamboyantengineer/ui-components/core';
import { act, cleanup, render, renderHook, screen } from '@testing-library/react';
import { userEvent } from '@testing-library/user-event';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { invertWorkbookOperations } from '../model/operations.js';
import { requireDefined } from '../model/require-defined.js';
import { createWorkbook } from '../model/workbook.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { HistoryPanel, PropertiesPanel } from './review-panels.js';
import { useSpreadsheetEditor } from './use-editor.js';

afterEach(cleanup);

function Panels(props: SpreadsheetEditorProps) {
  const editor = useSpreadsheetEditor(props);
  return (
    <UiProvider>
      <HistoryPanel editor={editor} />
      <PropertiesPanel editor={editor} />
      <output aria-label="Panel feedback">{editor.notice}</output>
    </UiProvider>
  );
}

describe('SE-HOST-1.1: panel commands use accepted host capabilities', () => {
  it('refreshes an untouched title from the host while retaining a local title draft', async () => {
    const user = userEvent.setup();
    const workbook = { ...createWorkbook(), name: 'Initial title' };
    const { rerender } = render(<Panels value={workbook} />);
    const title = screen.getByRole('textbox', { name: 'Workbook title' }) as HTMLInputElement;
    rerender(<Panels value={{ ...workbook, name: 'Remote title' }} />);
    expect(title.value).toBe('Remote title');
    await user.clear(title);
    await user.type(title, 'Local draft');
    rerender(<Panels value={{ ...workbook, name: 'Later remote title' }} />);
    expect(title.value).toBe('Local draft');
  });

  it('retains a title draft and reports a synchronous host rejection without an uncaught UI error', async () => {
    const user = userEvent.setup();
    const onChange = vi.fn(() => {
      throw new Error('Title change rejected');
    });
    render(<Panels defaultValue={createWorkbook()} onChange={onChange} />);
    const title = screen.getByRole('textbox', { name: 'Workbook title' }) as HTMLInputElement;
    await user.clear(title);
    await user.type(title, 'Local draft');
    await user.click(screen.getByRole('button', { name: 'Update title' }));
    expect(title.value).toBe('Local draft');
    expect(screen.getByLabelText('Panel feedback').textContent).toBe('Title change rejected');
    expect(onChange).toHaveBeenCalledTimes(1);
  });

  it('introduces appended line identities before writing their cells and removes them on undo', () => {
    const workbook = createWorkbook([['original']]);
    workbook.worksheets[0].rowIds = ['row-original'];
    workbook.worksheets[0].columnIds = ['column-original'];
    const onChange = vi.fn();
    const { result } = renderHook(() => useSpreadsheetEditor({ defaultValue: workbook, onChange }));
    act(() => {
      result.current.proposeCells([['new']], 1, 1);
    });
    const [next, context] = onChange.mock.calls[0];
    const rowId = next.worksheets[0].rowIds[1];
    const columnId = next.worksheets[0].columnIds[1];
    expect(context.operations).toMatchObject([
      { kind: 'lines', axis: 'row', action: 'insert', index: 1, count: 1, ids: [rowId] },
      { kind: 'lines', axis: 'column', action: 'insert', index: 1, count: 1, ids: [columnId] },
      { kind: 'cells', changes: [{ address: 'B2', rowId, columnId }] },
    ]);
    act(() => {
      result.current.undo();
    });
    expect(onChange.mock.calls[1][1].operations).toMatchObject([
      { kind: 'cells', changes: [{ address: 'B2', after: { value: null } }] },
      { kind: 'lines', axis: 'column', action: 'delete', extentOnly: true, ids: [columnId] },
      { kind: 'lines', axis: 'row', action: 'delete', extentOnly: true, ids: [rowId] },
    ]);
    expect(result.current.workbook).toEqual(workbook);
  });
  it('keeps delegated undo and redo available after a remote value replaces local history', async () => {
    const user = userEvent.setup();
    const undo = vi.fn(),
      redo = vi.fn();
    const props: SpreadsheetEditorProps = {
      value: createWorkbook([['original']]),
      undo: { canUndo: true, canRedo: true, undo, redo },
    };
    const { rerender } = render(<Panels {...props} />);
    await user.click(screen.getByRole('button', { name: 'Undo last change' }));
    expect(undo).toHaveBeenCalledOnce();
    rerender(<Panels {...props} value={createWorkbook([['peer']])} />);
    await user.click(screen.getByRole('button', { name: 'Redo last change' }));
    expect(redo).toHaveBeenCalledOnce();
    rerender(
      <Panels
        {...props}
        undo={{ ...requireDefined(props.undo), canUndo: false, canRedo: false }}
      />,
    );
    expect(screen.getByRole('button', { name: 'Undo last change' }).hasAttribute('disabled')).toBe(
      true,
    );
    expect(screen.getByRole('button', { name: 'Redo last change' }).hasAttribute('disabled')).toBe(
      true,
    );
  });

  it('publishes a workbook-title operation whose inverse removes a previously absent title', async () => {
    const user = userEvent.setup();
    const onChange = vi.fn();
    render(<Panels defaultValue={createWorkbook()} onChange={onChange} />);
    const title = screen.getByRole('textbox', { name: 'Workbook title' });
    await user.clear(title);
    await user.type(title, 'Review notes');
    await user.click(screen.getByRole('button', { name: 'Update title' }));
    const [workbook, context] = onChange.mock.calls[0];
    expect(workbook.name).toBe('Review notes');
    expect(context.operations).toEqual([
      { kind: 'workbook-metadata', before: { name: undefined }, after: { name: 'Review notes' } },
    ]);
    expect(invertWorkbookOperations(context.operations)).toEqual([
      { kind: 'workbook-metadata', before: { name: 'Review notes' }, after: { name: undefined } },
    ]);
    await act(async () => {
      await user.click(screen.getByRole('button', { name: 'Undo last change' }));
    });
    expect(onChange.mock.calls[1][0].name).toBeUndefined();
    expect(onChange.mock.calls[1][1].operations).toEqual(
      invertWorkbookOperations(context.operations),
    );
  });
});
