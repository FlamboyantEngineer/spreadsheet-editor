import { calculationCellKey } from '../calculation/publication.js';
import { createDocumentInsertion } from '../document/insertion.js';
import { selectionAddress, selectionCells } from '../model/editing.js';
import { assertWorkbook, copyWorkbook, type Workbook } from '../model/workbook.js';
import type { EditorCore } from './editor-core.js';
import { replaceWorkbookChange } from './editor-transactions.js';

/** Refuses a document table while any selected cell lacks a current ready result. */
function assertReadyTable(core: EditorCore) {
  const { props, sheet, selection, display } = core;
  const { calculation, readyCells } = display;
  if (!calculation.result || display.draftCell)
    throw new Error('A document table needs current calculation results.');
  for (const { x, y } of selectionCells(selection)) {
    const input = sheet.data[y]?.[x],
      key = calculationCellKey(sheet.id, x, y);
    const explicit =
      props.calculation?.mode === 'external'
        ? props.calculation.snapshot?.cellStates?.[key]
        : undefined;
    if (
      !readyCells.has(key) &&
      ((input != null && input !== '') || (explicit ?? calculation.status) !== 'ready')
    )
      throw new Error('Every inserted cell must have a current ready result.');
  }
}

/** Hands the selection to the host's document as a table or a live spreadsheet. */
export function insertSelection(core: EditorCore, kind: 'table' | 'spreadsheet') {
  if (kind === 'table') assertReadyTable(core);
  const { props, workbook, sheet, selection, display } = core;
  props.onInsert?.(
    createDocumentInsertion(workbook, display.calculation.result, {
      mode: kind,
      sheetId: sheet.id,
      range: selectionAddress(selection),
      display: props.display,
    }),
  );
}

/** Asks the host to refresh formula data, then recalculates. */
export async function refreshData(core: EditorCore, selected = false) {
  const { props, sheet, selection } = core;
  if (!core.canEdit || !props.onRefresh) return;
  await props.onRefresh(
    selected ? { sheetId: sheet.id, range: selectionAddress(selection) } : undefined,
  );
  if (core.live.current.canEdit) core.setRefreshVersion((version) => version + 1);
}

/** Replaces the workbook with one the host imported, unless it changed meanwhile. */
export async function importExcel(core: EditorCore) {
  const { props } = core;
  if (!core.canEdit || !props.onImportXlsx) return;
  const snapshot = core.workbook;
  const imported = await props.onImportXlsx();
  if (!imported) return;
  if (!core.alive.current || !core.live.current.canEdit) return;
  if (core.latestWorkbook.current !== snapshot)
    throw new Error('The workbook changed while importing. Try again.');
  assertWorkbook(imported);
  replaceWorkbookChange(core, imported, { label: 'Import Excel workbook', reason: 'import' });
}

/** Whether every authored cell of `workbook` has a ready result. */
function allReady(workbook: Workbook, readyCells: ReadonlySet<string>) {
  return workbook.worksheets.every((target) =>
    target.data.every((row, y) =>
      row.every(
        (input, x) =>
          input == null || input === '' || readyCells.has(calculationCellKey(target.id, x, y)),
      ),
    ),
  );
}

/** Hands the workbook and its current results to the host's Excel export. */
export async function exportExcel(core: EditorCore) {
  const { props, workbook, display } = core;
  const { calculation } = display;
  if (!props.onExportXlsx || !calculation.result || display.draftCell) return;
  if (!allReady(workbook, calculation.readyCells))
    throw new Error('Excel export requires current ready results for every authored cell.');
  await props.onExportXlsx(copyWorkbook(workbook), structuredClone(calculation.result));
}
