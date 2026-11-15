import type { MenuEntry } from '@flamboyantengineer/ui-components/menu';
import { requireDefined } from '../model/require-defined.js';
import { copyWorkbook, createWorkbook } from '../model/workbook.js';
import { alignmentGlyphs, editorCommandMetadata } from './editor-commands.js';
import {
  appendLine,
  cellContextMenu,
  columnContextMenu,
  type MenuContext,
  menuContext,
  rowContextMenu,
} from './editor-menu-items.js';
import type { SpreadsheetEditorController } from './use-editor.js';

function editMenu({ e, item, write }: MenuContext): MenuEntry[] {
  return [
    item('undo', 'Undo', { action: e.undo, isDisabled: write || !e.canUndo }),
    item('redo', 'Redo', { action: e.redo, isDisabled: write || !e.canRedo }),
    { id: 'clipboard-separator', type: 'separator' },
    item('cut', 'Cut', { action: () => e.copy(true), isDisabled: write }),
    item('copy', 'Copy', { action: () => e.copy(), isDisabled: false }),
    item('paste', 'Paste', { action: e.paste, isDisabled: write }),
    item('clear', 'Clear contents', { action: () => e.clear(), isDisabled: write }),
    item('clear-format', 'Clear formatting', { action: () => e.clear(true), isDisabled: write }),
    { id: 'search-separator', type: 'separator' },
    item('find', 'Find and replace', {
      action: () => e.setShowFind(true),
      isDisabled: e.props.features?.search === false,
    }),
  ];
}

function insertMenu({ e, item, structureBlocked }: MenuContext): MenuEntry[] {
  return [
    item('row-above', 'Insert rows above', {
      action: () => e.structure('row', 'insert'),
      isDisabled: structureBlocked,
    }),
    item('row-below', 'Insert rows below', {
      action: () => e.structure('row', 'insert', true),
      isDisabled: structureBlocked,
    }),
    item('column-before', 'Insert columns before', {
      action: () => e.structure('column', 'insert'),
      isDisabled: structureBlocked,
    }),
    item('column-after', 'Insert columns after', {
      action: () => e.structure('column', 'insert', true),
      isDisabled: structureBlocked,
    }),
    item('add-sheet', 'Add worksheet', {
      action: () => e.sheetCommand('add'),
      isDisabled: structureBlocked,
    }),
    { id: 'comment-separator', type: 'separator' },
    item('comment', 'Add comment', { action: e.beginComment, isDisabled: !e.canComment }),
  ];
}

function alignMenu({ e, write }: MenuContext): MenuEntry[] {
  return (['top', 'middle', 'bottom'] as const).flatMap((vertical) =>
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
}

function bordersMenu({ e, item, write }: MenuContext): MenuEntry[] {
  return (['top', 'left', 'all', 'right', 'bottom', 'none', 'outside'] as const).map((kind) =>
    item(
      `borders-${kind}`,
      kind === 'none' ? 'No borders' : `${kind[0].toUpperCase()}${kind.slice(1)} borders`,
      { action: () => e.border(kind), isDisabled: write },
    ),
  );
}

function formatMenu(
  context: MenuContext,
  { align, borders }: { align: MenuEntry[]; borders: MenuEntry[] },
): MenuEntry[] {
  const { e, item, write, structureBlocked } = context;
  const single = e.selection.x1 === e.selection.x2 && e.selection.y1 === e.selection.y2;
  return [
    item('bold', 'Bold', {
      action: () => e.toggleFormat('font-weight', 'bold'),
      isDisabled: write,
    }),
    item('italic', 'Italic', {
      action: () => e.toggleFormat('font-style', 'italic'),
      isDisabled: write,
    }),
    item('strike', 'Strikethrough', {
      action: () => e.toggleFormat('text-decoration', 'line-through'),
      isDisabled: write,
    }),
    { id: 'align', label: 'Alignment', ...editorCommandMetadata('align'), items: align },
    { id: 'borders', label: 'Borders', ...editorCommandMetadata('borders'), items: borders },
    item('wrap', 'Wrap text', {
      action: () => e.toggleFormat('white-space', 'pre-wrap'),
      isDisabled: write,
    }),
    item('merge', 'Merge cells', {
      action: () => e.merge(),
      isDisabled: structureBlocked || single,
    }),
    item('unmerge', 'Unmerge cells', {
      action: () => e.merge(true),
      isDisabled: structureBlocked || !Object.keys(e.sheet.mergeCells ?? {}).length,
    }),
    item('column-width', 'Column width…', {
      action: () => e.openDialog('column-width'),
      isDisabled: structureBlocked,
    }),
    item('row-height', 'Row height…', {
      action: () => e.openDialog('row-height'),
      isDisabled: structureBlocked,
    }),
  ];
}

/** The file menu's workbook commands: new, import and export. */
function workbookFileItems({ e, item, structureBlocked }: MenuContext): MenuEntry[] {
  const features = e.props.features;
  return [
    ...(features?.newWorkbook === false
      ? []
      : [
          item('new', 'New workbook', {
            action: () => e.replaceWorkbook(createWorkbook(), 'New workbook', 'new-workbook'),
            isDisabled: structureBlocked,
          }),
        ]),
    ...(features?.import === false
      ? []
      : [
          item('import', 'Import CSV or workbook JSON…', {
            action: () => e.openDialog('import'),
            isDisabled: structureBlocked,
          }),
        ]),
    item('export', 'Export workbook JSON…', { action: () => e.openDialog('export') }),
    ...(e.props.onImportXlsx && features?.import !== false
      ? [
          item('import-excel', 'Import Excel…', {
            action: e.importExcel,
            isDisabled: structureBlocked,
          }),
        ]
      : []),
    ...(e.props.onExportXlsx
      ? [
          item('export-excel', 'Export to Excel', {
            action: e.exportExcel,
            isDisabled: e.draftDirty || e.calculation.pending || !!e.calculation.error,
          }),
        ]
      : []),
  ];
}

/** The file menu's host commands: save, refresh and insertion into a document. */
function hostFileItems({ e, item, structureBlocked }: MenuContext): MenuEntry[] {
  return [
    ...(e.props.session
      ? [
          item('save', 'Save', {
            action: () => e.save(),
            isDisabled: !e.canEdit || e.savePhase === 'saving',
          }),
        ]
      : []),
    ...(e.props.onRefresh
      ? [
          item('refresh', 'Update all formula data', {
            action: () => e.refresh(),
            isDisabled: structureBlocked,
          }),
        ]
      : []),
    ...(e.props.onInsert
      ? [
          { id: 'insert-separator', type: 'separator' } as MenuEntry,
          item('insert-table', 'Insert as document table', {
            action: () => e.insert('table'),
            isDisabled: e.calculation.pending || !!e.calculation.error,
          }),
          item('insert-live', 'Insert as live spreadsheet', {
            action: () => e.insert('spreadsheet'),
          }),
        ]
      : []),
  ];
}

/** Sets the active worksheet's frozen column count. */
function freezeColumns(e: SpreadsheetEditorController, count: number, label: string) {
  return e.mutateWithDraft((base) => {
    const next = copyWorkbook(base);
    requireDefined(next.worksheets.find((sheet) => sheet.id === e.sheet.id)).freezeColumns = count;
    return next;
  }, label);
}

function viewMenu({ e, item, structureBlocked }: MenuContext): MenuEntry[] {
  return [
    {
      ...item('gridlines', 'Gridlines', { action: () => e.setShowGridlines(!e.showGridlines) }),
      isSelected: e.showGridlines,
    },
    {
      ...item('formulas', 'Show formulas', { action: () => e.setShowFormulas(!e.showFormulas) }),
      isSelected: e.showFormulas,
    },
    item('freeze', 'Freeze through selected column', {
      action: () => freezeColumns(e, e.selection.x1 + 1, 'Freeze columns'),
      isDisabled: structureBlocked,
    }),
    item('unfreeze', 'Unfreeze columns', {
      action: () => freezeColumns(e, 0, 'Unfreeze columns'),
      isDisabled: structureBlocked || !e.sheet.freezeColumns,
    }),
    { id: 'pane-separator', type: 'separator' },
    ...['properties', 'history', 'comments', 'suggestions'].map((id) =>
      item(`pane-${id}`, id[0].toUpperCase() + id.slice(1), { action: () => e.openPanel(id) }),
    ),
    ...(e.props.panels?.map((panel) =>
      item(`pane-${panel.id}`, panel.label, { action: () => e.openPanel(panel.id) }),
    ) ?? []),
  ];
}

function reviewMenu({ e, item }: MenuContext): MenuEntry[] {
  return [
    item('author', 'Authoring mode', {
      action: () => e.setMode('author'),
      isDisabled: e.locked || !e.capabilities.edit,
    }),
    item('review', 'Reviewing mode', {
      action: () => e.setMode('review'),
      isDisabled: e.locked || !e.capabilities.suggest || e.props.features?.suggestions === false,
    }),
    item('read', 'Reading mode', { action: () => e.setMode('read'), isDisabled: e.locked }),
    { id: 'review-separator', type: 'separator' },
    item('new-comment', 'Add comment', { action: e.beginComment, isDisabled: !e.canComment }),
    item('comments', 'Comments', { action: () => e.openPanel('comments') }),
    item('suggestions', 'Suggestions', { action: () => e.openPanel('suggestions') }),
  ];
}

export function editorMenus(e: SpreadsheetEditorController) {
  const context = menuContext(e);
  const edit = editMenu(context),
    insert = insertMenu(context),
    align = alignMenu(context),
    borders = bordersMenu(context);
  const format = formatMenu(context, { align, borders });
  return {
    file: [...workbookFileItems(context), ...hostFileItems(context)],
    edit,
    insert,
    format,
    view: viewMenu(context),
    review: reviewMenu(context),
    align,
    borders,
    context: cellContextMenu(context, { edit, insert, format }),
    rowContext: rowContextMenu(context),
    columnContext: columnContextMenu(context),
    appendRow: () => appendLine(e, 'row'),
    appendColumn: () => appendLine(e, 'column'),
  };
}
