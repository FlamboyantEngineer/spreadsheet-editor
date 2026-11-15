import { invertWorkbookOperations, type WorkbookOperation } from '../model/operations.js';
import { requireDefined } from '../model/require-defined.js';
import { parseSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { copyWorkbook, type Workbook, type Worksheet } from '../model/workbook.js';

const CELL_MAPS = [
  ['style', 'style'],
  ['formats', 'format'],
  ['inputValues', 'inputValue'],
  ['cellProvenance', 'provenance'],
] as const;

/** The parts of a worksheet a cell write can change besides its cells, to restore them exactly. */
interface SheetShape {
  rows: number;
  rowLengths: Map<number, number>;
  present: Set<string>;
}
/** Optional sheet maps a step may create empty; restored to absent when they were absent. */
const OPTIONAL_MAPS = ['style', 'formats', 'inputValues', 'cellProvenance', 'mergeCells'] as const;
type WorkbookShape = Map<string, SheetShape>;

/**
 * One local undo step (SE-9). A step made only of operations that carry their complete before
 * and after states keeps those operations and the shape of both sides; undo and redo rebuild the
 * workbook from them. Any other step (lines, sheets, replacement) keeps a snapshot of the side
 * it returns to.
 */
type HistoryEntry = { label: string; operations: WorkbookOperation[] } & (
  { shapes: { before: WorkbookShape; after: WorkbookShape } } | { snapshot: Workbook }
);

const REBUILT = new Set<WorkbookOperation['kind']>([
  'cells',
  'geometry',
  'merges',
  'annotations',
  'workbook-metadata',
  'sheet-rename',
  'sheet-reorder',
]);

function shapeOf(workbook: Workbook, operations: readonly WorkbookOperation[]): WorkbookShape {
  const shapes: WorkbookShape = new Map();
  for (const operation of operations) {
    if (operation.kind !== 'cells' && operation.kind !== 'geometry' && operation.kind !== 'merges')
      continue;
    const sheet = requireDefined(
      workbook.worksheets.find((value) => value.id === operation.sheetId),
    );
    let shape = shapes.get(sheet.id);
    if (!shape) {
      shape = {
        rows: sheet.data.length,
        rowLengths: new Map(),
        present: new Set(OPTIONAL_MAPS.filter((field) => sheet[field])),
      };
      shapes.set(sheet.id, shape);
    }
    if (operation.kind !== 'cells') continue;
    for (const change of operation.changes) {
      const { y } = requireDefined(parseSpreadsheetAddress(change.address));
      shape.rowLengths.set(y, sheet.data[y]?.length ?? 0);
    }
  }
  return shapes;
}

/** Record a step: `previous` is the workbook before it and `next` the one after it. */
export function historyEntry(
  previous: Workbook,
  next: Workbook,
  label: string,
  operations: WorkbookOperation[],
): HistoryEntry {
  return operations.every((operation) => REBUILT.has(operation.kind))
    ? {
        label,
        operations,
        shapes: { before: shapeOf(previous, operations), after: shapeOf(next, operations) },
      }
    : { label, operations, snapshot: copyWorkbook(previous) };
}

function restoreShape(sheet: Worksheet, shape: SheetShape) {
  sheet.data.length = Math.min(sheet.data.length, shape.rows);
  while (sheet.data.length < shape.rows) sheet.data.push([]);
  for (const [y, length] of shape.rowLengths) {
    if (y >= shape.rows) continue;
    const row = sheet.data[y];
    row.length = Math.min(row.length, length);
    while (row.length < length) row.push(null);
  }
  for (const field of OPTIONAL_MAPS)
    if (shape.present.has(field)) sheet[field] ??= {};
    else if (sheet[field] && !Object.keys(sheet[field]).length) delete sheet[field];
}

/** Rebuild one side of a rebuilt entry from `current`, the workbook on the other side. */
function rebuildWorkbook(
  current: Workbook,
  entry: Extract<HistoryEntry, { shapes: unknown }>,
  side: 'before' | 'after',
): Workbook {
  const next = copyWorkbook(current),
    sheetOf = (id: string) => requireDefined(next.worksheets.find((sheet) => sheet.id === id));
  for (const operation of entry.operations)
    switch (operation.kind) {
      case 'cells': {
        const sheet = sheetOf(operation.sheetId);
        for (const change of operation.changes) {
          const state = change[side],
            { x, y } = requireDefined(parseSpreadsheetAddress(change.address));
          while (sheet.data.length <= y) sheet.data.push([]);
          while (sheet.data[y].length <= x) sheet.data[y].push(null);
          sheet.data[y][x] = structuredClone(state.value);
          for (const [field, key] of CELL_MAPS) {
            const value = state[key];
            if (value === undefined) {
              if (sheet[field]) delete sheet[field][change.address];
            } else (sheet[field] ??= {} as never)[change.address] = structuredClone(value as never);
          }
        }
        break;
      }
      case 'geometry': {
        const sheet = sheetOf(operation.sheetId),
          geometry = operation[side];
        for (const field of ['minDimensions', 'columns', 'rows', 'freezeColumns'] as const)
          if (geometry[field] === undefined) delete sheet[field];
          else
            (sheet as unknown as Record<string, unknown>)[field] = structuredClone(geometry[field]);
        break;
      }
      case 'merges': {
        const sheet = sheetOf(operation.sheetId),
          merges = operation[side];
        sheet.mergeCells = structuredClone(merges);
        break;
      }
      case 'annotations':
        for (const field of ['comments', 'suggestions'] as const) {
          const value = operation[side][field];
          if (value === undefined) delete next[field];
          else next[field] = structuredClone(value) as never;
        }
        break;
      case 'workbook-metadata':
        if (operation[side].name === undefined) delete next.name;
        else next.name = operation[side].name;
        break;
      case 'sheet-rename':
        sheetOf(operation.sheetId).name = operation[side];
        break;
      case 'sheet-reorder':
        next.worksheets = operation[side].map(sheetOf);
        break;
      case 'lines':
      case 'sheet-add':
      case 'sheet-delete':
      case 'replace':
        // An entry holding any of these keeps a whole snapshot instead (REBUILT).
        throw new Error(`A ${operation.kind} step is restored from its snapshot.`);
    }
  for (const [id, shape] of entry.shapes[side]) restoreShape(sheetOf(id), shape);
  return next;
}

/** Local undo history; `head` is the workbook it belongs to. */
export interface EditorHistory {
  head: Workbook | null;
  past: HistoryEntry[];
  future: HistoryEntry[];
}

/** The workbook one step back, the operations that reach it, and the history after the step. */
export function undoStep(history: EditorHistory, workbook: Workbook) {
  const previous = history.past.at(-1);
  if (!previous) return null;
  const target =
    'snapshot' in previous ? previous.snapshot : rebuildWorkbook(workbook, previous, 'before');
  return {
    target,
    operations: invertWorkbookOperations(previous.operations),
    history: {
      head: target,
      past: history.past.slice(0, -1),
      future: [
        'snapshot' in previous ? { ...previous, snapshot: copyWorkbook(workbook) } : previous,
        ...history.future,
      ],
    },
  };
}

/** The workbook one step forward, the operations that reach it, and the history after the step. */
export function redoStep(history: EditorHistory, workbook: Workbook) {
  const next = history.future[0];
  if (!next) return null;
  const target = 'snapshot' in next ? next.snapshot : rebuildWorkbook(workbook, next, 'after');
  return {
    target,
    operations: next.operations,
    history: {
      head: target,
      past: [
        ...history.past,
        'snapshot' in next ? { ...next, snapshot: copyWorkbook(workbook) } : next,
      ],
      future: history.future.slice(1),
    },
  };
}
