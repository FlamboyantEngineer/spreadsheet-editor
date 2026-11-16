import { formatWorksheetCell } from '../calculation/cell-format.js';
import { calculationCellKey } from '../calculation/publication.js';
import { createSpreadsheetCopy, pasteSpreadsheetCopy } from '../model/clipboard.js';
import { parseDelimited, serializeDelimited } from '../model/delimited.js';
import { readSelection, writeCells } from '../model/editing.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { CellInput, Workbook } from '../model/workbook.js';
import type { SpreadsheetClipboard } from './contracts.js';
import { blockTargets } from './edit-targets.js';
import type { DraftText, EditorCore } from './editor-core.js';
import { mutateWithDraftChange } from './editor-transactions.js';

/** The host's clipboard, or the browser's. */
function editorClipboard(clipboard: SpreadsheetClipboard | undefined): SpreadsheetClipboard {
  return (
    clipboard ?? {
      async readText() {
        if (!globalThis.navigator?.clipboard)
          throw new Error('Clipboard access is unavailable. Use the grid’s paste shortcut.');
        return globalThis.navigator.clipboard.readText();
      },
      async writeText(text: string) {
        if (!globalThis.navigator?.clipboard) throw new Error('Clipboard access is unavailable.');
        await globalThis.navigator.clipboard.writeText(text);
      },
    }
  );
}

/** The selection as the reader sees it: shown values, and the draft in its cell. */
function displayedSelection(
  core: EditorCore,
  source: Workbook,
  draft: NonNullable<DraftText> | null,
) {
  const { sheet, selection, props, display } = core;
  return readSelection(source, sheet.id, selection).map((row, dy) =>
    row.map((cell, dx) => {
      const x = selection.x1 + dx,
        y = selection.y1 + dy;
      if (draft && x === selection.x1 && y === selection.y1) return draft.value;
      if (!display.readyCells.has(calculationCellKey(sheet.id, x, y))) return cell;
      return formatWorksheetCell(
        display.shown.result?.worksheets.find((item) => item.id === sheet.id)?.values[y]?.[x] ??
          null,
        sheet.formats?.[formatSpreadsheetAddress(x, y)],
        props.display,
      );
    }),
  );
}

/** Copies the selection's shown values, and keeps its cells for a paste back into the workbook. */
export async function copyCells(core: EditorCore, cut = false) {
  if (cut && !core.canWrite) return;
  const { sheet, selection, props, display, draftRef, copied } = core;
  const pending = draftRef.current,
    hasDraft = pending?.key === display.draftKey && pending.value !== String(display.authored),
    source = hasDraft
      ? writeCells(display.previewWorkbook, [[pending.value]], {
          sheetId: sheet.id,
          x: selection.x1,
          y: selection.y1,
          idFactory: props.idFactory,
        })
      : display.previewWorkbook;
  const text = serializeDelimited(displayedSelection(core, source, hasDraft ? pending : null));
  await editorClipboard(props.clipboard).writeText(text);
  const range = { sheetId: sheet.id, range: selection };
  if (cut && draftRef.current !== pending) {
    copied.current = createSpreadsheetCopy(source, range, { text, cut: false });
    core.warn('The cell edit changed while copying. Nothing was cut. Try Cut again.');
    return;
  }
  copied.current = createSpreadsheetCopy(source, range, { text, cut });
}

/** Pastes `values` at `{x, y}`: the editor's own copy keeps its formulas and formats. */
export function pasteValues(
  core: EditorCore,
  values: CellInput[][],
  { x, y, text: originalText }: { x: number; y: number; text?: string | undefined },
) {
  if (!core.canWrite) return;
  const { sheet, selection, props, copied } = core;
  const text = originalText ?? serializeDelimited(values);
  const cached = copied.current,
    matching = cached && cached.text.replace(/\r\n/g, '\n') === text.replace(/\r\n/g, '\n');
  const changed = mutateWithDraftChange(
    core,
    (base) =>
      matching
        ? pasteSpreadsheetCopy(base, cached, {
            sheetId: sheet.id,
            x,
            y,
            destination: selection,
            idFactory: props.idFactory,
          })
        : writeCells(base, values, { sheetId: sheet.id, x, y, idFactory: props.idFactory }),
    {
      label: 'Paste cells',
      written: matching ? undefined : blockTargets(sheet.id, { x, y }, values),
    },
  );
  if (changed && cached?.cut && matching) copied.current = null;
}

export async function pasteFromClipboard(core: EditorCore) {
  if (!core.canWrite) return;
  const text = await editorClipboard(core.props.clipboard).readText();
  pasteValues(core, parseDelimited(text), { x: core.selection.x1, y: core.selection.y1, text });
}
