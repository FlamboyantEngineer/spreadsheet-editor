import type { KeyboardEvent } from 'react';
import type { SpreadsheetEditorController } from './use-editor.js';

type Action = (() => void | Promise<void>) | undefined;

const isModifier = (event: KeyboardEvent) => event.ctrlKey || event.metaKey;

/** Shortcuts that work anywhere in the editor: find, save and a new comment. */
function editorShortcut(e: SpreadsheetEditorController, event: KeyboardEvent, key: string): Action {
  const modifier = isModifier(event);
  if (modifier && event.altKey && key === 'm') return e.beginComment;
  if (modifier && key === 's' && e.props.session) return () => e.save();
  if (modifier && key === 'f' && e.props.features?.search !== false)
    return () => e.setShowFind(true);
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
  key === 'enter' ||
  key === 'delete' ||
  key === 'backspace';

function isTextTarget(target: EventTarget) {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    (target instanceof HTMLElement && target.isContentEditable)
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
  openCellMenu: () => void,
) {
  if (event.nativeEvent.isComposing || event.repeat) return;
  if (event.altKey && !(isModifier(event) && event.key.toLowerCase() === 'm')) return;
  const key = event.key.toLowerCase();
  let action = editorShortcut(e, event, key);
  const gridTarget = event.target instanceof Element && !!event.target.closest('.se-editor__grid');
  if (!isTextTarget(event.target) && gridTarget) {
    if (e.lockedSelection && writesCell(event, key)) {
      claim(event);
      e.warn(`${e.lockedSelection.name} is editing a cell in this selection.`);
      return;
    }
    action = gridShortcut(e, event, { key, openCellMenu }) ?? action;
  }
  if (action) {
    claim(event);
    e.attempt(action);
  }
}
