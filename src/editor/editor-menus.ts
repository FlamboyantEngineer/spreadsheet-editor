import type { MenuEntry, MenuItemEntry } from '@flamboyantengineer/ui-components/menu';
import type { SpreadsheetEditorController } from './use-editor.js';
import { changeWorkbookStructure } from '../model/structure.js';
import { findOpenComment } from '../model/review.js';
import { selectionAddress } from '../model/editing.js';
import { copyWorkbook, createWorkbook } from '../model/workbook.js';

export function editorMenus(editor: SpreadsheetEditorController) {
  const e = editor,
    write = !e.canWrite,
    structureBlocked = !e.canEdit;
  const item = (
    id: string,
    label: string,
    action: () => void | Promise<void>,
    isDisabled = false,
    shortcut?: string,
  ): MenuItemEntry => ({ id, label, isDisabled, shortcut, onAction: () => e.attempt(action) });
  const edit: MenuEntry[] = [
    item('undo', 'Undo', e.undo, write || !e.history.past.length, 'Ctrl+Z'),
    item('redo', 'Redo', e.redo, write || !e.history.future.length, 'Ctrl+Y'),
    { id: 'clipboard-separator', type: 'separator' },
    item('cut', 'Cut', () => e.copy(true), write, 'Ctrl+X'),
    item('copy', 'Copy', () => e.copy(), false, 'Ctrl+C'),
    item('paste', 'Paste', e.paste, write, 'Ctrl+V'),
    item('clear', 'Clear contents', () => e.clear(), write),
    item('clear-format', 'Clear formatting', () => e.clear(true), write),
    { id: 'search-separator', type: 'separator' },
    item('find', 'Find and replace', () => e.setShowFind(true), false, 'Ctrl+F'),
  ];
  const insert: MenuEntry[] = [
    item('row-above', 'Insert rows above', () => e.structure('row', 'insert'), structureBlocked),
    item(
      'row-below',
      'Insert rows below',
      () => e.structure('row', 'insert', true),
      structureBlocked,
    ),
    item(
      'column-before',
      'Insert columns before',
      () => e.structure('column', 'insert'),
      structureBlocked,
    ),
    item(
      'column-after',
      'Insert columns after',
      () => e.structure('column', 'insert', true),
      structureBlocked,
    ),
    item('add-sheet', 'Add worksheet', () => e.sheetCommand('add'), structureBlocked),
    { id: 'comment-separator', type: 'separator' },
    item('comment', 'Add comment', e.beginComment, !e.canComment, 'Ctrl+Alt+M'),
  ];
  const align: MenuEntry[] = (['top', 'middle', 'bottom'] as const).flatMap((vertical) =>
    (['left', 'center', 'right'] as const).map((horizontal) => ({
      id: `${vertical}-${horizontal}`,
      label: `${vertical[0].toUpperCase()}${vertical.slice(1)} ${horizontal}`,
      isDisabled: write,
      isSelected: e.style['text-align'] === horizontal && e.style['vertical-align'] === vertical,
      onAction: () =>
        e.attempt(() => e.format({ 'text-align': horizontal, 'vertical-align': vertical })),
    })),
  );
  const borders: MenuEntry[] = (
    ['top', 'left', 'all', 'right', 'bottom', 'none', 'outside'] as const
  ).map((kind) =>
    item(
      `borders-${kind}`,
      kind === 'none' ? 'No borders' : `${kind[0].toUpperCase()}${kind.slice(1)} borders`,
      () => e.border(kind),
      write,
    ),
  );
  const format: MenuEntry[] = [
    item('bold', 'Bold', () => e.toggleFormat('font-weight', 'bold'), write, 'Ctrl+B'),
    item('italic', 'Italic', () => e.toggleFormat('font-style', 'italic'), write, 'Ctrl+I'),
    item('strike', 'Strikethrough', () => e.toggleFormat('text-decoration', 'line-through'), write),
    { id: 'align', label: 'Alignment', items: align },
    { id: 'borders', label: 'Borders', items: borders },
    item('wrap', 'Wrap text', () => e.toggleFormat('white-space', 'pre-wrap'), write),
    item(
      'merge',
      'Merge cells',
      () => e.merge(),
      structureBlocked || (e.selection.x1 === e.selection.x2 && e.selection.y1 === e.selection.y2),
    ),
    item(
      'unmerge',
      'Unmerge cells',
      () => e.merge(true),
      structureBlocked || !Object.keys(e.sheet.mergeCells ?? {}).length,
    ),
    item('column-width', 'Column width…', () => e.openDialog('column-width'), structureBlocked),
    item('row-height', 'Row height…', () => e.openDialog('row-height'), structureBlocked),
  ];
  const file: MenuEntry[] = [
    item('new', 'New workbook', () => e.mutate(createWorkbook(), 'New workbook'), structureBlocked),
    item('import', 'Import CSV or workbook JSON…', () => e.openDialog('import'), structureBlocked),
    item('export', 'Export workbook JSON…', () => e.openDialog('export')),
    ...(e.props.onImportXlsx
      ? [item('import-excel', 'Import Excel…', e.importExcel, structureBlocked)]
      : []),
    ...(e.props.onExportXlsx
      ? [
          item(
            'export-excel',
            'Export to Excel',
            e.exportExcel,
            e.draftDirty || e.calculation.pending || !!e.calculation.error,
          ),
        ]
      : []),
    ...(e.props.onSave
      ? [item('save', 'Save', () => e.save(), !e.canEdit || e.savePhase === 'saving', 'Ctrl+S')]
      : []),
    ...(e.props.onRefresh
      ? [item('refresh', 'Update all formula data', () => e.refresh(), structureBlocked)]
      : []),
    ...(e.props.onInsert
      ? [
          { id: 'insert-separator', type: 'separator' } as MenuEntry,
          item(
            'insert-table',
            'Insert as document table',
            () => e.insert('table'),
            e.calculation.pending || !!e.calculation.error,
          ),
          item('insert-live', 'Insert as live spreadsheet', () => e.insert('spreadsheet')),
        ]
      : []),
  ];
  const view: MenuEntry[] = [
    {
      ...item('gridlines', 'Gridlines', () => e.setShowGridlines(!e.showGridlines)),
      isSelected: e.showGridlines,
    },
    {
      ...item('formulas', 'Show formulas', () => e.setShowFormulas(!e.showFormulas)),
      isSelected: e.showFormulas,
    },
    item(
      'freeze',
      'Freeze through selected column',
      () => {
        const next = copyWorkbook(e.workbook);
        next.worksheets.find((sheet) => sheet.id === e.sheet.id)!.freezeColumns =
          e.selection.x1 + 1;
        e.mutate(next, 'Freeze columns');
      },
      structureBlocked,
    ),
    item(
      'unfreeze',
      'Unfreeze columns',
      () => {
        const next = copyWorkbook(e.workbook);
        next.worksheets.find((sheet) => sheet.id === e.sheet.id)!.freezeColumns = 0;
        e.mutate(next, 'Unfreeze columns');
      },
      structureBlocked || !e.sheet.freezeColumns,
    ),
    { id: 'pane-separator', type: 'separator' },
    ...['properties', 'history', 'comments', 'suggestions'].map((id) =>
      item(`pane-${id}`, id[0].toUpperCase() + id.slice(1), () => e.openPanel(id)),
    ),
    ...(e.props.panels?.map((panel) =>
      item(`pane-${panel.id}`, panel.label, () => e.openPanel(panel.id)),
    ) ?? []),
  ];
  const review: MenuEntry[] = [
    item(
      'author',
      'Authoring mode',
      () => e.setMode('author'),
      e.props.readOnly || e.props.permissions?.edit === false,
    ),
    item(
      'review',
      'Reviewing mode',
      () => e.setMode('review'),
      e.props.readOnly || e.props.permissions?.suggest === false,
    ),
    item('read', 'Reading mode', () => e.setMode('read'), e.props.readOnly),
    { id: 'review-separator', type: 'separator' },
    item('new-comment', 'Add comment', e.beginComment, !e.canComment),
    item('comments', 'Comments', () => e.openPanel('comments')),
    item('suggestions', 'Suggestions', () => e.openPanel('suggestions')),
  ];
  const context: MenuEntry[] = [
    ...edit.slice(2, 8),
    { id: 'context-insert', type: 'separator' },
    ...insert.slice(0, 4),
    item(
      'delete-rows',
      'Delete selected rows',
      () => e.structure('row', 'delete'),
      structureBlocked,
    ),
    item(
      'delete-columns',
      'Delete selected columns',
      () => e.structure('column', 'delete'),
      structureBlocked,
    ),
    { id: 'context-format', type: 'separator' },
    { id: 'format', label: 'Format cells', items: format },
    item(
      'context-comment',
      findOpenComment(e.workbook, e.sheet.id, selectionAddress(e.selection))
        ? 'Open comment'
        : 'Add comment',
      e.beginComment,
      !e.canComment,
    ),
    ...(e.props.onRefresh
      ? [
          item('refresh-selection', 'Update formula data', () => e.refresh(true), structureBlocked),
          item('refresh-all', 'Update all formula data', () => e.refresh(), structureBlocked),
        ]
      : []),
    item('context-suggestions', 'Show suggestions', () => e.openPanel('suggestions')),
  ];
  const rowContext: MenuEntry[] = [
    item('row-above', 'Insert rows above', () => e.structure('row', 'insert'), structureBlocked),
    item(
      'row-below',
      'Insert rows below',
      () => e.structure('row', 'insert', true),
      structureBlocked,
    ),
    item(
      'row-delete',
      'Delete selected rows',
      () => e.structure('row', 'delete'),
      structureBlocked,
    ),
    item('row-height', 'Row height…', () => e.openDialog('row-height'), structureBlocked),
  ];
  const columnContext: MenuEntry[] = [
    item(
      'col-before',
      'Insert columns before',
      () => e.structure('column', 'insert'),
      structureBlocked,
    ),
    item(
      'col-after',
      'Insert columns after',
      () => e.structure('column', 'insert', true),
      structureBlocked,
    ),
    item(
      'col-delete',
      'Delete selected columns',
      () => e.structure('column', 'delete'),
      structureBlocked,
    ),
    item('column-width', 'Column width…', () => e.openDialog('column-width'), structureBlocked),
  ];
  const sheets: MenuEntry[] = [
    item(
      'insert-sheet',
      'Insert worksheet before',
      () => e.sheetCommand('insert'),
      structureBlocked,
    ),
    item('rename', 'Rename worksheet…', () => e.openDialog('rename-sheet'), structureBlocked),
    item('duplicate', 'Duplicate worksheet', () => e.sheetCommand('duplicate'), structureBlocked),
    item(
      'move-left',
      'Move worksheet left',
      () => e.sheetCommand('left'),
      structureBlocked || e.workbook.worksheets.length === 1,
    ),
    item(
      'move-right',
      'Move worksheet right',
      () => e.sheetCommand('right'),
      structureBlocked || e.workbook.worksheets.length === 1,
    ),
    item(
      'delete',
      'Delete worksheet',
      () => e.sheetCommand('delete'),
      structureBlocked || e.workbook.worksheets.length === 1,
    ),
  ];
  return {
    file,
    edit,
    insert,
    format,
    view,
    review,
    align,
    borders,
    context,
    rowContext,
    columnContext,
    sheets,
    appendRow: () => {
      const change = { axis: 'row', kind: 'insert', index: e.sheet.data.length, count: 1 } as const;
      e.mutate(changeWorkbookStructure(e.workbook, e.sheet.id, change), 'Add row', {
        sheetId: e.sheet.id,
        change,
      });
    },
    appendColumn: () => {
      const change = {
        axis: 'column',
        kind: 'insert',
        index: Math.max(1, ...e.sheet.data.map((row) => row.length)),
        count: 1,
      } as const;
      e.mutate(changeWorkbookStructure(e.workbook, e.sheet.id, change), 'Add column', {
        sheetId: e.sheet.id,
        change,
      });
    },
  };
}
