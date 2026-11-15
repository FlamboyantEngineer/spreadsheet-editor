'use client';

import { Button } from '@flamboyantengineer/ui-components/button';
import { Menubar, MenubarMenu } from '@flamboyantengineer/ui-components/menubar';
import {
  ToolbarRow,
  ToolbarRowGroup,
  ToolbarRowSeparator,
} from '@flamboyantengineer/ui-components/toolbar-row';
import { editorCommandMetadata, editorGlyphs } from './editor-commands.js';
import { NumberFormattingGroup, TextFormattingGroup } from './editor-format-groups.js';
import { editorMenus } from './editor-menus.js';
import { MenuButton, type ToolbarSectionProps } from './editor-toolbar-controls.js';
import { ToolbarHeader } from './editor-toolbar-header.js';
import type { SpreadsheetEditorController } from './use-editor.js';

const capitalized = (name: string) => name[0].toUpperCase() + name.slice(1);

/** The embedded chrome's single menu of every spreadsheet menu. */
function ActionsMenu({ menus }: ToolbarSectionProps) {
  return (
    <>
      <MenuButton
        label="Spreadsheet actions"
        variant="secondary"
        icon={editorGlyphs.actions}
        iconOnly
        items={(['file', 'edit', 'insert', 'view', 'review'] as const).map((id) => ({
          id,
          label: capitalized(id),
          ...editorCommandMetadata(id),
          items: menus[id],
        }))}
      />
      <ToolbarRowSeparator />
    </>
  );
}

function HistoryGroup({ e, disabled }: ToolbarSectionProps) {
  return (
    <ToolbarRowGroup className="se-editor__tool-group" aria-label="History">
      <Button
        size="sm"
        icon={editorGlyphs.undo}
        label="Undo"
        variant="ghost"
        onPress={e.undo}
        isDisabled={disabled || !e.canUndo}
        tooltip="Undo"
        shortcut="Ctrl+Z"
      />
      <Button
        size="sm"
        icon={editorGlyphs.redo}
        label="Redo"
        variant="ghost"
        onPress={e.redo}
        isDisabled={disabled || !e.canRedo}
        tooltip="Redo"
        shortcut="Ctrl+Y"
      />
    </ToolbarRowGroup>
  );
}

function CellLayoutGroup({ e, menus, disabled }: ToolbarSectionProps) {
  const single = e.selection.x1 === e.selection.x2 && e.selection.y1 === e.selection.y2;
  return (
    <ToolbarRowGroup className="se-editor__tool-group" aria-label="Cell layout">
      <MenuButton
        label="Align cells"
        icon={editorGlyphs.align}
        items={menus.align}
        disabled={disabled}
        iconOnly
      />
      <MenuButton
        label="Borders"
        icon={editorGlyphs.borders}
        items={menus.borders}
        disabled={disabled}
        iconOnly
      />
      <Button
        size="sm"
        icon={editorGlyphs.wrap}
        label="Wrap text"
        tooltip="Wrap text"
        isDisabled={disabled}
        isSelected={e.style['white-space'] === 'pre-wrap'}
        onPress={() => e.attempt(() => e.toggleFormat('white-space', 'pre-wrap'))}
      />
      <Button
        size="sm"
        icon={editorGlyphs.merge}
        label="Merge cells"
        tooltip="Merge cells"
        onPress={() => e.attempt(() => e.merge())}
        isDisabled={!e.canEdit || single}
      />
      <Button
        size="sm"
        icon={editorGlyphs.unmerge}
        label="Unmerge cells"
        tooltip="Unmerge cells"
        onPress={() => e.attempt(() => e.merge(true))}
        isDisabled={!e.canEdit || !Object.keys(e.sheet.mergeCells ?? {}).length}
      />
    </ToolbarRowGroup>
  );
}

/** The comments and suggestions panel toggles, with their open and pending counts. */
function ReviewButtons({ e }: ToolbarSectionProps) {
  const open = e.workbook.comments?.filter((thread) => thread.state === 'open').length ?? 0;
  const pending = e.workbook.suggestions?.filter((item) => item.status === 'pending').length ?? 0;
  return (
    <>
      <Button
        size="sm"
        icon={editorCommandMetadata('context-comment').icon}
        label={`Comments (${open})`}
        tooltip="Comments"
        onPress={() => e.openPanel('comments')}
        isSelected={e.panelOpen && e.panel === 'comments'}
      />
      <Button
        size="sm"
        icon={editorGlyphs.review}
        label={`Suggestions (${pending})`}
        tooltip="Suggestions"
        onPress={() => e.openPanel('suggestions')}
        isSelected={e.panelOpen && e.panel === 'suggestions'}
      />
    </>
  );
}

function WorkbookCommandsGroup(props: ToolbarSectionProps) {
  const { e, menus } = props;
  return (
    <>
      <ToolbarRowSeparator />
      <ToolbarRowGroup className="se-editor__tool-group" aria-label="Workbook commands">
        <Button
          size="sm"
          icon={editorCommandMetadata('row-below').icon}
          label="Add row"
          tooltip="Add row"
          onPress={() => e.attempt(menus.appendRow)}
          isDisabled={!e.canEdit}
        />
        <Button
          size="sm"
          icon={editorCommandMetadata('column-after').icon}
          label="Add column"
          tooltip="Add column"
          onPress={() => e.attempt(menus.appendColumn)}
          isDisabled={!e.canEdit}
        />
        <Button
          size="sm"
          icon={editorCommandMetadata('find').icon}
          label="Find"
          tooltip="Find"
          isDisabled={e.props.features?.search === false}
          onPress={() => e.setShowFind(!e.showFind)}
          isSelected={e.showFind}
        />
        <ReviewButtons {...props} />
        {e.props.onRefresh && (
          <Button
            size="sm"
            icon={editorCommandMetadata('refresh').icon}
            label="Refresh data"
            tooltip="Refresh data"
            isDisabled={!e.canEdit}
            onPress={() => e.attempt(() => e.refresh())}
          />
        )}
      </ToolbarRowGroup>
    </>
  );
}

export function EditorToolbar({ editor: e }: { editor: SpreadsheetEditorController }) {
  const menus = editorMenus(e);
  const chrome = e.props.chrome ?? 'full';
  const embedded = chrome === 'embedded';
  if (chrome === 'minimal') return null;
  const section: ToolbarSectionProps = { e, menus, disabled: !e.canWrite, embedded };
  return (
    <div className="se-editor__chrome" onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
      <ToolbarHeader {...section} />
      {!embedded && (
        <Menubar aria-label="Spreadsheet menu" className="se-editor__menubar">
          {(['file', 'edit', 'insert', 'format', 'view', 'review'] as const).map((name) => (
            <MenubarMenu key={name} label={capitalized(name)} items={menus[name]} />
          ))}
        </Menubar>
      )}
      <ToolbarRow
        aria-label="Spreadsheet formatting"
        className="se-editor__toolbar"
        overflow="scroll"
        justify="start"
      >
        {embedded && <ActionsMenu {...section} />}
        <HistoryGroup {...section} />
        <ToolbarRowSeparator />
        <TextFormattingGroup {...section} />
        <ToolbarRowSeparator />
        <NumberFormattingGroup {...section} />
        <ToolbarRowSeparator />
        <CellLayoutGroup {...section} />
        {!embedded && <WorkbookCommandsGroup {...section} />}
      </ToolbarRow>
    </div>
  );
}
