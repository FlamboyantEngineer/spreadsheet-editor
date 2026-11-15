// @vitest-environment jsdom
import { act, cleanup, fireEvent, render, screen, waitFor } from '@testing-library/react';
import { createRef, useState } from 'react';
import { afterAll, afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { requireDefined } from '../model/require-defined.js';
import {
  EditorFormulaBar,
  type EditorFormulaBarHandle,
  type EditorFormulaBarProps,
} from './editor-formula-bar.js';

afterEach(cleanup);
// React Aria queries generated metadata names through CSS.escape; jsdom omits that browser API.
beforeAll(() =>
  vi.stubGlobal('CSS', { escape: (value: string) => value.replace(/[^\w-]/g, '\\$&') }),
);
afterAll(() => vi.unstubAllGlobals());

function configuration(overrides: Partial<EditorFormulaBarProps> = {}): EditorFormulaBarProps {
  return {
    selection: { x1: 0, x2: 0, y1: 0, y2: 0 },
    draftValue: '=A2+B3',
    result: '3',
    onDraftChange: vi.fn(),
    onCommit: vi.fn(),
    onCancel: vi.fn(),
    onNavigate: vi.fn(),
    ...overrides,
  };
}

function ControlledBar(props: EditorFormulaBarProps) {
  const [draftValue, setDraftValue] = useState(props.draftValue);
  return (
    <EditorFormulaBar
      {...props}
      draftValue={draftValue}
      onDraftChange={(value) => {
        setDraftValue(value);
        props.onDraftChange(value);
      }}
    />
  );
}

function formula() {
  return screen.getByRole('textbox', { name: 'Cell value or formula' }) as HTMLTextAreaElement;
}

function address() {
  return screen.getByRole('textbox', { name: 'Selected cell or range' }) as HTMLInputElement;
}

describe('SE-CHROME-4.1: public formula bar editing surface', () => {
  it('uses the library code surface and semantic reference tokens without field actions', () => {
    const { container } = render(<EditorFormulaBar {...configuration()} />);
    expect(formula().tagName).toBe('TEXTAREA');
    expect(formula().rows).toBe(1);
    expect(formula().value).toBe('=A2+B3');
    expect(screen.getByRole('group', { name: 'Formula bar' }).contains(formula())).toBe(true);
    expect(container.querySelector('[data-token="reference-0"]')?.textContent).toBe('A2');
    expect(container.querySelector('[data-token="reference-1"]')?.textContent).toBe('B3');
    expect(screen.queryByRole('button', { name: /apply|cancel|restore|reset/i })).toBeNull();
    fireEvent.change(address(), { target: { value: 'b2' } });
    expect(screen.queryByRole('button', { name: /apply|cancel|restore|reset/i })).toBeNull();
  });

  it('delegates an unhandled Enter once and Escape once, including CodeEditor prevented Enter', () => {
    const props = configuration();
    render(<EditorFormulaBar {...props} />);
    fireEvent.keyDown(formula(), { key: 'Enter' });
    expect(props.onCommit).toHaveBeenCalledExactlyOnceWith('down');
    expect(props.onNavigate).not.toHaveBeenCalled();
    fireEvent.keyDown(formula(), { key: 'Escape' });
    expect(props.onCancel).toHaveBeenCalledTimes(1);
    expect(props.onDraftChange).not.toHaveBeenCalled();
  });

  it('delegates Control/Command+Enter as a commit without navigation', () => {
    const props = configuration();
    render(<EditorFormulaBar {...props} />);
    fireEvent.keyDown(formula(), { key: 'Enter', ctrlKey: true });
    fireEvent.keyDown(formula(), { key: 'Enter', metaKey: true });
    expect(props.onCommit).toHaveBeenNthCalledWith(1, 'stay');
    expect(props.onCommit).toHaveBeenNthCalledWith(2, 'stay');
    expect(props.onNavigate).not.toHaveBeenCalled();
  });

  it('leaves IME keys to the input method, including composition state and candidate key 229', () => {
    const props = configuration();
    render(<EditorFormulaBar {...props} />);
    fireEvent.compositionStart(formula());
    fireEvent.keyDown(formula(), { key: 'Enter' });
    fireEvent.keyDown(formula(), { key: 'Escape' });
    fireEvent.compositionEnd(formula());
    fireEvent.keyDown(formula(), { key: 'Enter', isComposing: true });
    fireEvent.keyDown(formula(), { key: 'Escape', keyCode: 229 });
    expect(props.onCommit).not.toHaveBeenCalled();
    expect(props.onCancel).not.toHaveBeenCalled();
    fireEvent.keyDown(formula(), { key: 'Enter' });
    expect(props.onCommit).toHaveBeenCalledExactlyOnceWith('down');
  });

  it('lets completion acceptance use Enter before committing the formula', async () => {
    const props = configuration({ draftValue: '' });
    render(<ControlledBar {...props} />);
    act(() => formula().focus());
    fireEvent.change(formula(), {
      target: { value: '=suMsq', selectionStart: 6, selectionEnd: 6 },
    });
    await waitFor(() =>
      expect(screen.queryByRole('option', { name: 'SUMSQ SUMSQ function' })).not.toBeNull(),
    );
    // The public list becomes actionable on its next paint, once the reader can see it.
    await act(() => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())));
    fireEvent.keyDown(formula(), { key: 'Enter' });
    await waitFor(() => expect(formula().value).toBe('=SUMSQ('));
    expect(props.onCommit).not.toHaveBeenCalled();
    fireEvent.keyDown(formula(), { key: 'Enter' });
    expect(props.onCommit).toHaveBeenCalledExactlyOnceWith('down');
  });

  it('dismisses completions before Escape cancels the edit', async () => {
    const props = configuration({ draftValue: '' });
    render(<ControlledBar {...props} />);
    act(() => formula().focus());
    fireEvent.change(formula(), {
      target: { value: '=suMsq', selectionStart: 6, selectionEnd: 6 },
    });
    await waitFor(() =>
      expect(screen.queryByRole('option', { name: 'SUMSQ SUMSQ function' })).not.toBeNull(),
    );
    fireEvent.keyDown(formula(), { key: 'Escape' });
    expect(props.onCancel).not.toHaveBeenCalled();
    fireEvent.keyDown(formula(), { key: 'Escape' });
    expect(props.onCancel).toHaveBeenCalledTimes(1);
  });

  it('inserts a picked range at the caret and keeps the parent-owned draft focused', () => {
    const ref = createRef<EditorFormulaBarHandle>();
    const props = configuration({ ref, draftValue: '=SUM(A1)+1' });
    render(<ControlledBar {...props} />);
    act(() => {
      formula().focus();
      formula().setSelectionRange(5, 7);
    });
    act(() => requireDefined(ref.current).insertReference('B2:C3'));
    expect(props.onDraftChange).toHaveBeenCalledExactlyOnceWith('=SUM(B2:C3)+1');
    expect(formula().value).toBe('=SUM(B2:C3)+1');
    expect(formula().selectionStart).toBe(10);
    expect(formula().selectionEnd).toBe(10);
    expect(document.activeElement).toBe(formula());
    expect(props.onCommit).not.toHaveBeenCalled();
  });

  it('restores the insertion caret when a controlled host publishes the draft later', () => {
    const ref = createRef<EditorFormulaBarHandle>();
    const props = configuration({ ref, draftValue: '=A2+B3' });
    const view = render(<EditorFormulaBar {...props} />);
    act(() => {
      formula().focus();
      formula().setSelectionRange(1, 3);
      requireDefined(ref.current).insertReference('C4:D5');
    });
    expect(formula().value).toBe('=A2+B3');
    expect(props.onDraftChange).toHaveBeenCalledExactlyOnceWith('=C4:D5+B3');
    view.rerender(<EditorFormulaBar {...props} draftValue="=C4:D5+B3" />);
    expect(formula().selectionStart).toBe(6);
    expect(formula().selectionEnd).toBe(6);
    expect(document.activeElement).toBe(formula());
  });

  it('accepts bounded A1 ranges live, marks invalid input, and normalizes Enter/Escape/blur', () => {
    const props = configuration();
    render(<EditorFormulaBar {...props} />);
    fireEvent.change(address(), { target: { value: 'c4:b2' } });
    expect(props.onNavigate).toHaveBeenLastCalledWith({ x1: 1, x2: 2, y1: 1, y2: 3 });
    fireEvent.keyDown(address(), { key: 'Enter' });
    expect(address().value).toBe('B2:C4');
    vi.mocked(props.onNavigate).mockClear();
    fireEvent.change(address(), { target: { value: 'XFE1' } });
    expect(address().getAttribute('aria-invalid')).toBe('true');
    expect(props.onNavigate).not.toHaveBeenCalled();
    fireEvent.keyDown(address(), { key: 'Enter' });
    expect(address().getAttribute('aria-invalid')).toBe('true');
    fireEvent.keyDown(address(), { key: 'Escape' });
    expect(address().value).toBe('A1');
    expect(address().getAttribute('aria-invalid')).not.toBe('true');
    fireEvent.change(address(), { target: { value: 'broken' } });
    fireEvent.blur(address());
    expect(address().value).toBe('A1');
  });

  it('holds drafts within the bar and host editing scope, then notifies when leaving', () => {
    const outside = document.createElement('button');
    const hostControl = document.createElement('button');
    const props = configuration({
      onLeaveEditor: vi.fn(),
      onEditingChange: vi.fn(),
      keepsEditing: (target) => target === hostControl,
    });
    render(<EditorFormulaBar {...props} />);
    fireEvent.focus(formula());
    expect(props.onEditingChange).toHaveBeenLastCalledWith(true);
    fireEvent.blur(formula(), { relatedTarget: address() });
    fireEvent.blur(formula(), { relatedTarget: hostControl });
    expect(props.onEditingChange).toHaveBeenLastCalledWith(false);
    fireEvent.focus(formula());
    expect(props.onEditingChange).toHaveBeenLastCalledWith(true);
    expect(props.onLeaveEditor).not.toHaveBeenCalled();
    fireEvent.blur(formula(), { relatedTarget: outside });
    expect(props.onLeaveEditor).toHaveBeenCalledTimes(1);
    expect(props.onEditingChange).toHaveBeenLastCalledWith(false);
  });

  it('shows current result status and leaves read-only raw formulas selectable', () => {
    const ref = createRef<EditorFormulaBarHandle>();
    const props = configuration({ ref, isReadOnly: true });
    const view = render(<EditorFormulaBar {...props} />);
    expect(formula().readOnly).toBe(true);
    act(() => {
      formula().focus();
      formula().setSelectionRange(1, 3);
      requireDefined(ref.current).insertReference('C5');
    });
    expect(formula().selectionStart).toBe(1);
    expect(formula().selectionEnd).toBe(3);
    fireEvent.keyDown(formula(), { key: 'Enter' });
    expect(props.onCommit).not.toHaveBeenCalled();
    expect(props.onDraftChange).not.toHaveBeenCalled();
    expect(screen.getByLabelText('Calculated cell value').textContent).toBe('3');
    view.rerender(<EditorFormulaBar {...props} pending />);
    expect(screen.getByLabelText('Calculated cell value').textContent).toBe('Calculating…');
    view.rerender(<EditorFormulaBar {...props} error />);
    expect(screen.getByLabelText('Calculated cell value').textContent).toBe('Unavailable');
    view.rerender(<EditorFormulaBar {...props} draftDirty />);
    expect(screen.getByLabelText('Calculated cell value').textContent).toBe('Editing…');
  });
});
