'use client';

import { useMemo } from 'react';
import { parseDelimited } from '../model/delimited.js';
import { writeCells } from '../model/editing.js';
import { selectionLock } from '../model/presence.js';
import { requireDefined } from '../model/require-defined.js';
import { type CellInput, createWorkbook, type Worksheet } from '../model/workbook.js';
import type { GridContext, WorkbookGridProps } from './grid-context.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import { captureWorksheetCellChanges, gridWorksheetExtent, sameCapture } from './worksheet-data.js';

type CellChange = { x: number; y: number; value: CellInput | undefined };
type ConfigContext = Pick<
  GridContext,
  'latest' | 'handledNativeChange' | 'suppress' | 'inCell' | 'applied' | 'setRevision'
>;

/** Whether an engine change may become a workbook change now. */
function capturing({ latest, suppress }: ConfigContext, allowReview: boolean): boolean {
  const current = latest.current;
  return !(
    suppress.current ||
    current.readOnly ||
    (current.reviewMode && !allowReview && !current.onReviewChange)
  );
}

/** The worksheet after the engine's changes, or after its whole data when it names none. */
function capturedWorksheet(
  context: ConfigContext,
  engine: JSpreadsheetWorksheetInstance,
  changes: readonly CellChange[] | undefined,
): Worksheet | undefined {
  const current = context.latest.current.worksheet;
  const data = changes ? undefined : engine?.getData?.();
  if (!changes && !data) return undefined;
  const merge = changes ? undefined : engine.getMerge?.();
  return {
    ...(changes
      ? captureWorksheetCellChanges(current, changes, context.latest.current.idFactory)
      : {
          ...current,
          data: requireDefined(data).map((row) =>
            row.map((cell) => (cell === undefined ? null : (cell as CellInput))),
          ),
        }),
    ...(merge && !Array.isArray(merge) ? { mergeCells: structuredClone(merge) } : {}),
  };
}

/**
 * A batch has already changed CE's presentation. Reconcile it with the
 * accepted workbook after host preflight refuses the complete command.
 */
function refuseCapture(context: ConfigContext, error: unknown) {
  const current = context.latest.current;
  if (current.onError) current.onError(error);
  else
    current.onBlockedAction?.(
      error instanceof Error ? error.message : 'The cell changes could not be applied.',
    );
  context.applied.current = null;
  context.setRevision((value) => value + 1);
}

function capture(
  context: ConfigContext,
  engine: unknown,
  { allowReview = false, changes }: { allowReview?: boolean; changes?: readonly CellChange[] },
) {
  if (context.handledNativeChange.current) {
    context.handledNativeChange.current = false;
    return;
  }
  if (!capturing(context, allowReview)) return;
  const current = context.latest.current.worksheet;
  try {
    const next = capturedWorksheet(context, engine as JSpreadsheetWorksheetInstance, changes);
    // Native data edits must never capture zoomed presentation measures as authored sizes.
    if (next && !sameCapture(current, next)) context.latest.current.onChange(next);
  } catch (error) {
    refuseCapture(context, error);
  }
}

/** Sends cell values to the host's review, or writes them to the worksheet. */
function propose(context: ConfigContext, values: CellInput[][], at: { x: number; y: number }) {
  const latest = context.latest.current;
  if (latest.onReviewChange) {
    latest.onReviewChange(values, at.x, at.y);
    return;
  }
  const current = latest.worksheet,
    workbook = createWorkbook();
  workbook.worksheets = [structuredClone(current)];
  latest.onChange(
    writeCells(workbook, values, {
      sheetId: current.id,
      x: Number(at.x),
      y: Number(at.y),
      idFactory: latest.idFactory,
    }).worksheets[0],
  );
}

const authoredAt = (current: WorkbookGridProps, x: number, y: number) =>
  current.worksheet.data[Number(y)]?.[Number(x)];

/** The peer whose edit lock covers a cell, if one does. */
function lockedBy(current: WorkbookGridProps, x: number, y: number) {
  return selectionLock(
    { schemaVersion: 1, worksheets: [current.worksheet] },
    {
      sheetId: current.worksheet.id,
      selection: { x1: Number(x), x2: Number(x), y1: Number(y), y2: Number(y) },
    },
    current.collaborators ?? [],
  );
}

/**
 * An edit in the cell's own editor: unchanged text keeps its typed value, and in review the
 * host gets the edit before CE mutates the cell or discards its editor.
 */
function editedValue(
  context: ConfigContext,
  { x, y, value }: { x: number; y: number; value: CellInput },
): CellInput | undefined {
  const current = context.latest.current;
  const authored = authoredAt(current, x, y);
  if (context.inCell.current && String(authored ?? '') === String(value ?? '')) {
    // Text inputs stringify typed primitives. Finishing an unchanged edit
    // must preserve its logical/number type and any canonical input value.
    context.handledNativeChange.current = true;
    return authored ?? '';
  }
  if (!context.inCell.current || (!current.reviewMode && !current.onReviewChange)) return undefined;
  // Host preflight runs before CE mutates the cell or discards its editor.
  // A synchronous rejection therefore retains the exact active native draft.
  propose(context, [[value]], { x: Number(x), y: Number(y) });
  context.handledNativeChange.current = true;
  const after = context.latest.current;
  return after.reviewMode ? (authoredAt(after, x, y) ?? '') : undefined;
}

/**
 * The engine's `onbeforechange`: the value to write instead, or undefined to let the engine
 * write it. Vendor arguments: (worksheet, cell, x, y, value).
 */
function beforeChange(context: ConfigContext, args: readonly unknown[]): CellInput | undefined {
  const [, , x, y, value] = args as [unknown, unknown, number, number, CellInput];
  const current = context.latest.current;
  if (current.readOnly) return authoredAt(current, x, y) ?? '';
  const locked = lockedBy(current, x, y);
  if (locked && !context.suppress.current) {
    current.onBlockedAction?.(`${locked.name} is editing this cell.`);
    return authoredAt(current, x, y) ?? '';
  }
  if (context.suppress.current) return undefined;
  return editedValue(context, { x, y, value });
}

/** Pasted engine data as cell inputs: delimited text, or rows of values or cell objects. */
function pastedValues(data: unknown): CellInput[][] {
  if (typeof data === 'string') return parseDelimited(data);
  if (!Array.isArray(data)) return [];
  return data.map((row: unknown) =>
    Array.isArray(row)
      ? row.map((cell: unknown) =>
          cell && typeof cell === 'object' && 'value' in cell
            ? (cell.value as CellInput)
            : (cell as CellInput),
        )
      : [],
  );
}

/** The engine's `onbeforepaste`. Vendor arguments: (worksheet, data, x, y). */
function beforePaste(context: ConfigContext, args: readonly unknown[]): false | undefined {
  const [, data, x, y] = args as [unknown, unknown, number, number];
  const current = context.latest.current;
  if (current.readOnly) return false;
  if (!current.reviewMode && !current.onPasteCells) return undefined;
  const values = pastedValues(data);
  if (current.onPasteCells) current.onPasteCells(values, Number(x), Number(y));
  else propose(context, values, { x, y });
  return false;
}

/** The worksheet the engine mounts with; later changes reach it through the sync effects. */
function mountedWorksheet(props: WorkbookGridProps) {
  const { worksheet } = props;
  return {
    worksheetName: worksheet.name,
    data: worksheet.data.map((row) => [...row]),
    columns: worksheet.columns?.map((column) => ({ ...column, type: 'text' })),
    mergeCells: structuredClone(worksheet.mergeCells ?? {}),
    style: { ...worksheet.style },
    minDimensions: gridWorksheetExtent(worksheet, props.logicalExtent),
    columnResize: false,
    rowResize: false,
    rows: structuredClone(worksheet.rows ?? {}),
    freezeColumns: worksheet.freezeColumns ?? 0,
    editable: !props.readOnly,
    // React commands move every affected formula and review anchor together.
    allowInsertRow: false,
    allowDeleteRow: false,
    allowInsertColumn: false,
    allowDeleteColumn: false,
    tableOverflow: true,
    // The engine attaches only its first rows; the row window takes over once mounted.
    lazyLoading: true,
    tableWidth: '100%',
    tableHeight: '400px',
  };
}

/** The engine's configuration, made once for the grid's life. */
export function useGridConfig(context: ConfigContext, props: WorkbookGridProps) {
  return useMemo(
    () => ({
      about: false,
      toolbar: false,
      tabs: false,
      allowExport: false,
      parseFormulas: false,
      contextMenu: () => false,
      onbeforechange: (...args: unknown[]) => beforeChange(context, args),
      onbeforepaste: (...args: unknown[]) => beforePaste(context, args),
      worksheets: [mountedWorksheet(props)],
      onafterchanges: (worksheet: unknown, changes: CellChange[]) =>
        capture(context, worksheet, { changes }),
      onresizecolumn: (worksheet: unknown) => capture(context, worksheet, { allowReview: true }),
      onresizerow: (worksheet: unknown) => capture(context, worksheet, { allowReview: true }),
      // Canonical React history includes cell data, formatting and review state.
      onundo: () => false,
      onredo: () => false,
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps -- the engine is configured once; later props reach it through latest and the sync effects
    [],
  );
}
