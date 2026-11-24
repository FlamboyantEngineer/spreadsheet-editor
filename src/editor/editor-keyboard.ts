import type { KeyboardEvent } from 'react';
import type { SpreadsheetEditorController } from './use-editor.js';

type Action = (() => void | Promise<void>) | undefined;

const isModifier = (event: KeyboardEvent) => event.ctrlKey || event.metaKey;
const nativeClipboardShortcut = (
  e: SpreadsheetEditorController,
  event: KeyboardEvent,
  key: string,
) => !e.props.clipboard && isModifier(event) && ['c', 'x', 'v'].includes(key);

/** Shortcuts that work anywhere in the editor: find, save and a new comment. */
function editorShortcut(
  e: SpreadsheetEditorController,
  event: KeyboardEvent,
  { key, openFind }: { key: string; openFind: () => void },
): Action {
  const modifier = isModifier(event);
  if (modifier && event.altKey && key === 'm') return e.beginComment;
  if (modifier && key === 's' && e.props.session) return () => e.save();
  if (modifier && key === 'f' && e.props.features?.search !== false) return openFind;
  return undefined;
}

/** Clipboard, history and formatting shortcuts of the grid, and its context menu key. */
function gridShortcut(
  e: SpreadsheetEditorController,
  event: KeyboardEvent,
  { key, openCellMenu }: { key: string; openCellMenu: () => void },
): Action {
  const modifier = isModifier(event);
  if (key === 'contextmenu' || (key === 'f10' && event.shiftKey)) return openCellMenu;
  if (!modifier) return undefined;
  if (key === 'i' && e.canWrite) return () => e.toggleFormat('font-style', 'italic');
  if (key === 'b' && e.canWrite) return () => e.toggleFormat('font-weight', 'bold');
  if (key === 'y') return e.redo;
  if (key === 'z') return event.shiftKey ? e.redo : e.undo;
  return clipboardShortcut(e, key);
}

function clipboardShortcut(e: SpreadsheetEditorController, key: string): Action {
  if (key === 'c') return () => e.copy();
  if (key === 'x') return () => e.copy(true);
  if (key === 'v') return e.paste;
  return undefined;
}

/** Whether a key would write into the grid's selection. */
const writesCell = (event: KeyboardEvent, key: string) =>
  (key.length === 1 && !isModifier(event)) ||
  key === 'f2' ||
  key === 'delete' ||
  key === 'backspace';

function isGridTarget(target: EventTarget) {
  return (
    target instanceof Element &&
    !!target.closest('.se-editor__grid-surface') &&
    !target.closest('button, input, textarea, select, a[href], [contenteditable="true"]')
  );
}

function claim(event: KeyboardEvent) {
  event.preventDefault();
  event.stopPropagation();
}

/**
 * The editor's keyboard shortcuts (SE-CHROME). The grid's own shortcuts apply only while focus is
 * in the grid outside a text field; a collaborator's edit lock refuses keys that write.
 */
export function handleEditorKey(
  e: SpreadsheetEditorController,
  event: KeyboardEvent<HTMLElement>,
  { openCellMenu, openFind }: { openCellMenu: () => void; openFind: () => void },
) {
  if (event.nativeEvent.isComposing) return;
  if (event.altKey && !(isModifier(event) && event.key.toLowerCase() === 'm')) return;
  const key = event.key.toLowerCase();
  if (event.target instanceof Element && event.target.closest('.se-editor') !== event.currentTarget)
    return;
  let action = editorShortcut(e, event, { key, openFind });
  if (isGridTarget(event.target)) {
    if (e.lockedSelection && writesCell(event, key)) {
      claim(event);
      e.warn(`${e.lockedSelection.name} is editing a cell in this selection.`);
      return;
    }
    if (nativeClipboardShortcut(e, event, key)) {
      // CE handles these on document and cancels the browser copy/cut/paste event.
      // Stop CE without cancelling the default trusted clipboard event.
      event.stopPropagation();
      return;
    }
    action = gridShortcut(e, event, { key, openCellMenu }) ?? action;
  }
  if (action) {
    claim(event);
    if (!event.repeat) e.attempt(action);
  }
}
