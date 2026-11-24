import { formatWorksheetCell } from '../calculation/cell-format.js';
import { calculationCellKey } from '../calculation/publication.js';
import { createSpreadsheetCopy, pasteSpreadsheetCopy } from '../model/clipboard.js';
import { parseDelimited, serializeDelimited } from '../model/delimited.js';
import { writeCells } from '../model/editing.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import type { CellInput } from '../model/workbook.js';
import {
  type ClipboardCopy,
  clipboardHtml,
  readBrowserClipboard,
  readClipboardEvent,
  type SpreadsheetClipboardData,
  writeBrowserClipboard,
  writeClipboardEvent,
} from './clipboard-payload.js';
import { blockTargets } from './edit-targets.js';
import type { DraftText, EditorCore } from './editor-core.js';
import { mutateWithDraftChange } from './editor-transactions.js';

/** The selection as the reader sees it: shown values, and the draft in its cell. */
function displayedSelection(
  core: EditorCore,
  cells: CellInput[][],
  draft: NonNullable<DraftText> | null,
) {
  const { sheet, selection, props, display } = core;
  const output = display.shown.result?.worksheets.find((item) => item.id === sheet.id)?.values;
  return cells.map((row, dy) =>
    row.map((cell, dx) => {
      const x = selection.x1 + dx,
        y = selection.y1 + dy;
      const key = calculationCellKey(sheet.id, x, y);
      if (draft && x === selection.x1 && y === selection.y1) return draft.value;
      if (!display.readyCells.has(key)) return cell;
      const hint = display.presentation?.cells[key];
      const format = display.previewSheet.formats?.[formatSpreadsheetAddress(x, y)];
      if (hint?.text !== undefined) return hint.text;
      return formatWorksheetCell(
        output?.[y]?.[x] ?? null,
        format ?? (hint?.defaultUnit ? { kind: 'unit', unit: hint.defaultUnit } : undefined),
        props.display,
      );
    }),
  );
}

function prepareCopy(core: EditorCore, cut: boolean) {
  const { sheet, selection, props, display, draftRef } = core;
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
  const origin = globalThis.crypto.getRandomValues(new Uint32Array(4)).join('-');
  const copy = createSpreadsheetCopy(
    source,
    { sheetId: sheet.id, range: selection },
    { text: '', cut },
  );
  const values = displayedSelection(core, copy.cells, hasDraft ? pending : null);
  const text = serializeDelimited(values);
  return { pending, values, data: { text, origin }, cached: { copy: { ...copy, text }, origin } };
}

function browserCopy({ data, values }: ReturnType<typeof prepareCopy>) {
  return { text: data.text, html: clipboardHtml(data.origin, values) };
}

function mayFinishWrite(core: EditorCore) {
  return (
    core.alive.current && core.live.current.canWrite && core.live.current.canEdit === core.canEdit
  );
}

function rememberCopy(
  core: EditorCore,
  { pending, cached }: { pending: DraftText; cached: ClipboardCopy },
  keptOrigin: boolean,
) {
  if (!core.alive.current) return;
  if (cached.copy.cut && core.draftRef.current !== pending) {
    cached.copy.cut = false;
    core.warn('The cell edit changed while copying. Nothing was cut. Try Cut again.');
  } else if (cached.copy.cut && !keptOrigin) {
    core.warn('The clipboard only retained plain text. The cells were copied.');
  } else if (cached.copy.cut && !mayFinishWrite(core)) {
    cached.copy.cut = false;
    core.warn('Editing permissions changed while copying. The cells were copied.');
  }
  core.copied.current = keptOrigin ? cached : null;
}

/** Copies through a host adapter or the asynchronous browser clipboard used by menus. */
export async function copyCells(core: EditorCore, cut = false) {
  if (cut && !core.canWrite) return;
  const prepared = prepareCopy(core, cut);
  const adapter = core.props.clipboard;
  let keptOrigin = false;
  if (adapter?.rich) {
    await adapter.rich.write(prepared.data);
    keptOrigin = true;
  } else if (adapter) await adapter.writeText(prepared.data.text);
  else keptOrigin = await writeBrowserClipboard(browserCopy(prepared));
  rememberCopy(core, prepared, keptOrigin);
}

/** The browser's trusted copy and cut events do not need asynchronous clipboard permission. */
export function copyToClipboardEvent(core: EditorCore, transfer: DataTransfer, cut: boolean) {
  if (cut && !core.canWrite) return;
  const prepared = prepareCopy(core, cut);
  writeClipboardEvent(transfer, browserCopy(prepared));
  rememberCopy(core, prepared, true);
}

/** Writes an external clipboard block through the normal draft and transaction path. */
export function pasteValues(
  core: EditorCore,
  values: CellInput[][],
  { x, y }: { x: number; y: number },
) {
  const { sheet, props } = core;
  if (!core.canWrite) return;
  mutateWithDraftChange(
    core,
    (base) => writeCells(base, values, { sheetId: sheet.id, x, y, idFactory: props.idFactory }),
    { label: 'Paste cells', written: blockTargets(sheet.id, { x, y }, values) },
  );
}

function pasteData(core: EditorCore, data: SpreadsheetClipboardData) {
  if (!mayFinishWrite(core)) return;
  const { sheet, selection, props, copied } = core;
  const at = { x: selection.x1, y: selection.y1 };
  const cached = copied.current;
  if (
    !cached ||
    cached.origin !== data.origin ||
    cached.copy.text.replace(/\r\n/g, '\n') !== data.text.replace(/\r\n/g, '\n')
  ) {
    pasteValues(core, parseDelimited(data.text), at);
    return;
  }
  const changed = mutateWithDraftChange(
    core,
    (base) =>
      pasteSpreadsheetCopy(base, cached.copy, {
        sheetId: sheet.id,
        ...at,
        destination: selection,
        idFactory: props.idFactory,
      }),
    { label: 'Paste cells' },
  );
  if (changed && cached.copy.cut) copied.current = null;
}

export function pasteFromClipboardEvent(core: EditorCore, transfer: DataTransfer) {
  if (!core.canWrite) return;
  const data = readClipboardEvent(transfer);
  if (data) pasteData(core, data);
}

export async function pasteFromClipboard(core: EditorCore) {
  if (!core.canWrite) return;
  const adapter = core.props.clipboard;
  const data = adapter?.rich
    ? await adapter.rich.read()
    : adapter
      ? { text: await adapter.readText() }
      : await readBrowserClipboard();
  if (data) pasteData(core, data);
}
