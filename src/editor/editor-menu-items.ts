import type { MenuEntry, MenuItemEntry } from '@flamboyantengineer/ui-components/menu';
import { selectionAddress } from '../model/editing.js';
import { materializeWorksheetExtent, worksheetExtent } from '../model/extent.js';
import { changeWorkbookStructureWithOperations } from '../model/operations.js';
import { requireDefined } from '../model/require-defined.js';
import { findOpenComment } from '../model/review.js';
import { editorCommandMetadata } from './editor-commands.js';
import type { SpreadsheetEditorController } from './use-editor.js';

/** What a menu command runs, and whether it is unavailable. */
export interface MenuCommand {
  action: () => unknown;
  isDisabled?: boolean | undefined;
}

/** Builds a command entry with its shared shortcut and icon metadata. */
export type MenuItemFactory = (id: string, label: string, command: MenuCommand) => MenuItemEntry;

/** The editor and the availability flags every menu reads. */
export interface MenuContext {
  e: SpreadsheetEditorController;
  item: MenuItemFactory;
  /** Cell writes are unavailable. */
  write: boolean;
  /** Structure changes are unavailable. */
  structureBlocked: boolean;
}

export function menuContext(e: SpreadsheetEditorController): MenuContext {
  const item: MenuItemFactory = (id, label, { action, isDisabled = false }) => ({
    id,
    label,
    'aria-label': label,
    ...editorCommandMetadata(id),
    isDisabled,
    onAction: () => e.attempt(action),
  });
  return { e, item, write: !e.canWrite, structureBlocked: !e.canEdit };
}

/** The cell context menu, which reuses the clipboard and line entries of the main menus. */
export function cellContextMenu(
  { e, item, structureBlocked }: MenuContext,
  { edit, insert, format }: { edit: MenuEntry[]; insert: MenuEntry[]; format: MenuEntry[] },
): MenuEntry[] {
  const hasComment = findOpenComment(e.workbook, e.sheet.id, selectionAddress(e.selection));
  return [
    ...edit.slice(2, 8),
    { id: 'context-insert', type: 'separator' },
    ...insert.slice(0, 4),
    item('delete-rows', 'Delete selected rows', {
      action: () => e.structure('row', 'delete'),
      isDisabled: structureBlocked,
    }),
    item('delete-columns', 'Delete selected columns', {
      action: () => e.structure('column', 'delete'),
      isDisabled: structureBlocked,
    }),
    { id: 'context-format', type: 'separator' },
    { id: 'format', label: 'Format cells', ...editorCommandMetadata('format'), items: format },
    item('context-comment', hasComment ? 'Open comment' : 'Add comment', {
      action: e.beginComment,
      isDisabled: !e.canComment,
    }),
    ...(e.props.onRefresh
      ? [
          item('refresh-selection', 'Update formula data', {
            action: () => e.refresh(true),
            isDisabled: structureBlocked,
          }),
          item('refresh-all', 'Update all formula data', {
            action: () => e.refresh(),
            isDisabled: structureBlocked,
          }),
        ]
      : []),
    item('context-suggestions', 'Show suggestions', { action: () => e.openPanel('suggestions') }),
  ];
}

/** The row header context menu. */
export function rowContextMenu({ e, item, structureBlocked }: MenuContext): MenuEntry[] {
  return [
    item('row-above', 'Insert rows above', {
      action: () => e.structure('row', 'insert'),
      isDisabled: structureBlocked,
    }),
    item('row-below', 'Insert rows below', {
      action: () => e.structure('row', 'insert', true),
      isDisabled: structureBlocked,
    }),
    item('row-delete', 'Delete selected rows', {
      action: () => e.structure('row', 'delete'),
      isDisabled: structureBlocked,
    }),
    item('row-height', 'Row height…', {
      action: () => e.openDialog('row-height'),
      isDisabled: structureBlocked,
    }),
  ];
}

/** The column header context menu. */
export function columnContextMenu({ e, item, structureBlocked }: MenuContext): MenuEntry[] {
  return [
    item('col-before', 'Insert columns before', {
      action: () => e.structure('column', 'insert'),
      isDisabled: structureBlocked,
    }),
    item('col-after', 'Insert columns after', {
      action: () => e.structure('column', 'insert', true),
      isDisabled: structureBlocked,
    }),
    item('col-delete', 'Delete selected columns', {
      action: () => e.structure('column', 'delete'),
      isDisabled: structureBlocked,
    }),
    item('column-width', 'Column width…', {
      action: () => e.openDialog('column-width'),
      isDisabled: structureBlocked,
    }),
  ];
}

/** Adds one row or column after the worksheet's materialized extent. */
export function appendLine(e: SpreadsheetEditorController, axis: 'row' | 'column') {
  const extent = worksheetExtent(
    requireDefined(
      materializeWorksheetExtent(e.workbook, e.sheet.id).worksheets.find(
        (sheet) => sheet.id === e.sheet.id,
      ),
    ),
  );
  const change = {
    axis,
    kind: 'insert',
    index: axis === 'row' ? extent.rows : extent.columns,
    count: 1,
  } as const;
  e.mutateWithDraft(
    (base) =>
      changeWorkbookStructureWithOperations(
        base,
        { sheetId: e.sheet.id, idFactory: e.props.idFactory },
        change,
      ),
    axis === 'row' ? 'Add row' : 'Add column',
    { sheetId: e.sheet.id, change },
  );
}
