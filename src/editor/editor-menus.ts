import type { MenuEntry, MenuItemEntry } from '@flamboyantengineer/ui-components/menu';
import type { SpreadsheetEditorController } from './use-editor.js';
import { changeWorkbookStructureWithOperations } from '../model/operations.js';
import { materializeWorksheetExtent, worksheetExtent } from '../model/extent.js';
import { findOpenComment } from '../model/review.js';
import { selectionAddress } from '../model/editing.js';
import { copyWorkbook, createWorkbook } from '../model/workbook.js';
import { alignmentGlyphs, editorCommandMetadata } from './editor-commands.js';

export function editorMenus(editor: SpreadsheetEditorController) {
  const e = editor,
    write = !e.canWrite,
    structureBlocked = !e.canEdit;
  const item = (
    id: string,
    label: string,
    action: () => unknown,
    isDisabled = false,
  ): MenuItemEntry => ({
    id,
    label,
    'aria-label': label,
    ...editorCommandMetadata(id),
    isDisabled,
    onAction: () => e.attempt(action),
  });
  const edit: MenuEntry[] = [
    item('undo', 'Undo', e.undo, write || !e.canUndo),
    item('redo', 'Redo', e.redo, write || !e.canRedo),
    { id: 'clipboard-separator', type: 'separator' },
    item('cut', 'Cut', () => e.copy(true), write),
    item('copy', 'Copy', () => e.copy(), false),
    item('paste', 'Paste', e.paste, write),
    item('clear', 'Clear contents', () => e.clear(), write),
    item('clear-format', 'Clear formatting', () => e.clear(true), write),
    { id: 'search-separator', type: 'separator' },
    item('find', 'Find and replace', () => e.setShowFind(true), e.props.features?.search === false),
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
    item('comment', 'Add comment', e.beginComment, !e.canComment),
  ];
  const align: MenuEntry[] = (['top', 'middle', 'bottom'] as const).flatMap((vertical) =>
    (['left', 'center', 'right'] as const).map((horizontal) => ({
      id: `${vertical}-${horizontal}`,
      label: `${vertical[0].toUpperCase()}${vertical.slice(1)} ${horizontal}`,
      icon: alignmentGlyphs[`${vertical}-${horizontal}`],
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
    item('bold', 'Bold', () => e.toggleFormat('font-weight', 'bold'), write),
    item('italic', 'Italic', () => e.toggleFormat('font-style', 'italic'), write),
    item('strike', 'Strikethrough', () => e.toggleFormat('text-decoration', 'line-through'), write),
    { id: 'align', label: 'Alignment', ...editorCommandMetadata('align'), items: align },
    { id: 'borders', label: 'Borders', ...editorCommandMetadata('borders'), items: borders },
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
    ...(e.props.features?.newWorkbook === false
      ? []
      : [
          item(
            'new',
            'New workbook',
            () => e.replaceWorkbook(createWorkbook(), 'New workbook', 'new-workbook'),
            structureBlocked,
          ),
        ]),
    ...(e.props.features?.import === false
      ? []
      : [
          item(
            'import',
            'Import CSV or workbook JSON…',
            () => e.openDialog('import'),
            structureBlocked,
          ),
        ]),
    item('export', 'Export workbook JSON…', () => e.openDialog('export')),
    ...(e.props.onImportXlsx && e.props.features?.import !== false
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
      ? [item('save', 'Save', () => e.save(), !e.canEdit || e.savePhase === 'saving')]
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
      () =>
        e.mutateWithDraft((base) => {
          const next = copyWorkbook(base);
          next.worksheets.find((sheet) => sheet.id === e.sheet.id)!.freezeColumns =
            e.selection.x1 + 1;
          return next;
        }, 'Freeze columns'),
      structureBlocked,
    ),
    item(
      'unfreeze',
      'Unfreeze columns',
      () =>
        e.mutateWithDraft((base) => {
          const next = copyWorkbook(base);
          next.worksheets.find((sheet) => sheet.id === e.sheet.id)!.freezeColumns = 0;
          return next;
        }, 'Unfreeze columns'),
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
    item('author', 'Authoring mode', () => e.setMode('author'), e.locked || !e.capabilities.edit),
    item(
      'review',
      'Reviewing mode',
      () => e.setMode('review'),
      e.locked || !e.capabilities.suggest || e.props.features?.suggestions === false,
    ),
    item('read', 'Reading mode', () => e.setMode('read'), e.locked),
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
    { id: 'format', label: 'Format cells', ...editorCommandMetadata('format'), items: format },
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
    appendRow: () => {
      const change = {
        axis: 'row',
        kind: 'insert',
        index: worksheetExtent(
          materializeWorksheetExtent(e.workbook, e.sheet.id).worksheets.find(
            (sheet) => sheet.id === e.sheet.id,
          )!,
        ).rows,
        count: 1,
      } as const;
      e.mutateWithDraft(
        (base) =>
          changeWorkbookStructureWithOperations(base, e.sheet.id, change, e.props.idFactory),
        'Add row',
        {
          sheetId: e.sheet.id,
          change,
        },
      );
    },
    appendColumn: () => {
      const change = {
        axis: 'column',
        kind: 'insert',
        index: worksheetExtent(
          materializeWorksheetExtent(e.workbook, e.sheet.id).worksheets.find(
            (sheet) => sheet.id === e.sheet.id,
          )!,
        ).columns,
        count: 1,
      } as const;
      e.mutateWithDraft(
        (base) =>
          changeWorkbookStructureWithOperations(base, e.sheet.id, change, e.props.idFactory),
        'Add column',
        {
          sheetId: e.sheet.id,
          change,
        },
      );
    },
  };
}
