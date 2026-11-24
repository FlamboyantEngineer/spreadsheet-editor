import {
  borderCells,
  clearSelection,
  formatCells,
  mergeSelection,
  selectionAddress,
  selectionCells,
  writeCells,
} from '../model/editing.js';
import { extendWorksheetIdentities } from '../model/extent.js';
import { changeWorkbookStructureWithOperations } from '../model/operations.js';
import { requireDefined } from '../model/require-defined.js';
import type { CellFormat } from '../model/review.js';
import { parseSpreadsheetRange } from '../model/spreadsheet-addressing.js';
import type { StructureChange } from '../model/structure.js';
import type { CellInput, Worksheet } from '../model/workbook.js';
import { blockTargets } from './edit-targets.js';
import type { EditorCore } from './editor-core.js';
import {
  attemptAction,
  mutateChange,
  mutateWithDraftChange,
  publishChange,
} from './editor-transactions.js';
import { redoStep, undoStep } from './history.js';

/** The grid reports the kind of change and the authored cells it knows it touched. */
export type WorksheetEdit = { kind: 'layout' } | { kind: 'cells'; addresses?: readonly string[] };

function withLayout(
  previous: Worksheet,
  next: Worksheet,
  idFactory: EditorCore['props']['idFactory'],
) {
  const updated = {
    ...previous,
    ...(previous.rowIds && { rowIds: previous.rowIds.slice() }),
    ...(previous.columnIds && { columnIds: previous.columnIds.slice() }),
  };
  for (const field of ['columns', 'rows', 'minDimensions', 'freezeColumns'] as const) {
    if (next[field] === undefined) delete updated[field];
    else Object.assign(updated, { [field]: next[field] });
  }
  extendWorksheetIdentities(updated, idFactory);
  return updated;
}

/** Applies a worksheet the grid changed: a layout change keeps the draft, a cell edit ends it. */
export function updateSheet(core: EditorCore, next: Worksheet, change?: WorksheetEdit) {
  const { workbook } = core;
  const previous = workbook.worksheets.find((value) => value.id === next.id);
  if (!previous) return false;
  const edit = change ?? (next.data === previous.data ? { kind: 'layout' } : { kind: 'cells' });
  if (edit.kind === 'layout')
    return mutateWithDraftChange(
      core,
      (base) => ({
        ...base,
        worksheets: base.worksheets.map((value) =>
          value.id === next.id ? withLayout(value, next, core.props.idFactory) : value,
        ),
      }),
      { label: 'Update worksheet layout' },
    );
  // Native cell commits already contain the native editor's protected draft.
  const changed = mutateChange(
    core,
    {
      ...workbook,
      worksheets: workbook.worksheets.map((sheet) => (sheet.id === next.id ? next : sheet)),
    },
    { label: 'Edit cells', targets: edit.addresses && new Map([[next.id, edit.addresses]]) },
  );
  if (changed) core.setDraft(null);
  return changed;
}

export function proposeCells(
  core: EditorCore,
  values: CellInput[][],
  { x, y }: { x: number; y: number },
) {
  const { workbook, sheet, props } = core;
  const changed = mutateChange(
    core,
    writeCells(workbook, values, { sheetId: sheet.id, x, y, idFactory: props.idFactory }),
    { label: 'Edit cells', targets: blockTargets(sheet.id, { x, y }, values) },
  );
  if (changed) core.setDraft(null);
  return changed;
}

export function formatSelection(
  core: EditorCore,
  properties: Record<string, string | null>,
  format?: CellFormat | null,
) {
  const { props, sheet, selection } = core;
  if (format && props.formatting?.kinds && !props.formatting.kinds.includes(format.kind))
    throw new Error('The host does not support this number format.');
  if (format?.decimals !== undefined && props.formatting?.decimalPlaces === false)
    throw new Error('The host does not support decimal-place overrides.');
  mutateWithDraftChange(
    core,
    (base) =>
      formatCells(
        base,
        { sheetId: sheet.id, range: selection, idFactory: props.idFactory },
        { properties, format },
      ),
    {
      label: `Format ${selectionAddress(selection)}`,
      written: new Map([[sheet.id, selectionCells(selection).map((cell) => cell.address)]]),
    },
  );
}

export type BorderKind = 'all' | 'outside' | 'none' | 'top' | 'right' | 'bottom' | 'left';

export function borderSelection(core: EditorCore, kind: BorderKind) {
  const { sheet, selection, props } = core;
  mutateWithDraftChange(
    core,
    (base) =>
      borderCells(base, { sheetId: sheet.id, range: selection, idFactory: props.idFactory }, kind),
    { label: `Set ${kind} borders` },
  );
}

export function clearCells(core: EditorCore, formatting = false) {
  const { sheet, selection } = core;
  return mutateWithDraftChange(
    core,
    (base) => clearSelection(base, { sheetId: sheet.id, range: selection }, formatting),
    { label: formatting ? 'Clear formatting' : 'Clear cells' },
  );
}

export function mergeCells(core: EditorCore, unmerge = false) {
  if (!core.canEdit) return;
  const { sheet, selection, props } = core;
  mutateWithDraftChange(
    core,
    (base) =>
      mergeSelection(
        base,
        { sheetId: sheet.id, range: selection, idFactory: props.idFactory },
        unmerge,
      ),
    { label: unmerge ? 'Unmerge cells' : 'Merge cells' },
  );
}

/** One step back or forward, through the host's undo manager when it has one. */
export function historyStep(core: EditorCore, direction: 'undo' | 'redo') {
  if (!core.canWrite) return;
  const host = core.props.undo;
  if (host) {
    attemptAction(core, () => (direction === 'undo' ? host.undo() : host.redo()));
    return;
  }
  const step =
    direction === 'undo'
      ? undoStep(core.currentHistory, core.workbook)
      : redoStep(core.currentHistory, core.workbook);
  if (!step) return;
  publishChange(core, step.target, { operations: step.operations, source: direction });
  core.setDraft(null);
  core.setHistory(step.history);
}

/** Whether a line change splits a merged range of the active worksheet. */
function crossesMerge(
  sheet: Worksheet,
  { axis, kind, index }: StructureChange,
  [first, last]: [number, number],
) {
  return Object.entries(sheet.mergeCells ?? {}).some(([address, [columns, rows]]) => {
    const range = requireDefined(parseSpreadsheetRange(address));
    const start = axis === 'row' ? range.start.y : range.start.x;
    const end = start + (axis === 'row' ? rows : columns) - 1;
    return kind === 'insert' ? index > start && index <= end : first <= end && last >= start;
  });
}

/** Inserts or deletes the selected lines, asking first when the change splits a merge. */
export function structureChange(
  core: EditorCore,
  { axis, kind }: Pick<StructureChange, 'axis' | 'kind'>,
  after = false,
) {
  if (!core.canEdit) {
    core.warn('Worksheet structure can be changed only in Authoring mode.');
    return;
  }
  const { sheet, selection, props } = core;
  const first = axis === 'row' ? selection.y1 : selection.x1,
    last = axis === 'row' ? selection.y2 : selection.x2;
  const change = { axis, kind, index: after ? last + 1 : first, count: last - first + 1 };
  const apply = () =>
    mutateWithDraftChange(
      core,
      (base) =>
        changeWorkbookStructureWithOperations(
          base,
          { sheetId: sheet.id, idFactory: props.idFactory },
          change,
        ),
      {
        label: `${kind === 'insert' ? 'Insert' : 'Delete'} ${axis}${last > first ? 's' : ''}`,
        operation: { sheetId: sheet.id, change },
      },
    );
  if (crossesMerge(sheet, change, [first, last])) {
    core.pendingStructure.current = apply;
    core.setDialog('confirm-structure');
  } else apply();
}
