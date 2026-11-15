'use client';

import { Button } from '@flamboyantengineer/ui-components/button';
import { DropdownMenu } from '@flamboyantengineer/ui-components/dropdown-menu';
import type { IconSource } from '@flamboyantengineer/ui-components/icon';
import {
  ToolbarRow,
  ToolbarRowGroup,
  ToolbarRowSeparator,
} from '@flamboyantengineer/ui-components/toolbar-row';
import type { MenuEntry } from '@flamboyantengineer/ui-components/menu';
import { Menubar, MenubarMenu } from '@flamboyantengineer/ui-components/menubar';
import { AvatarGroup } from '@flamboyantengineer/ui-components/avatar-group';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { SaveStatus } from '@flamboyantengineer/ui-components/save-status';
import { Tooltip } from '@flamboyantengineer/ui-components/tooltip';
import { collaboratorColorIndex } from '../model/presence.js';
import type { SpreadsheetEditorController } from './use-editor.js';
import { UNIT_PRESETS } from '../calculation/units.js';
import { editorCommandMetadata, editorGlyphs, numberFormatGlyphs } from './editor-commands.js';
import { editorMenus } from './editor-menus.js';
import { EditorColorControl, EditorFontSize } from './text-format-controls.js';

function MenuButton({
  label,
  items,
  icon,
  disabled,
  iconOnly = false,
  hint,
  variant,
}: {
  label: string;
  items: MenuEntry[];
  icon: IconSource;
  disabled?: boolean;
  iconOnly?: boolean;
  hint?: string;
  variant?: 'secondary';
}) {
  return (
    <Tooltip content={hint ?? label} isDisabled={disabled}>
      <DropdownMenu
        label={label}
        aria-label={label}
        size="sm"
        variant={variant}
        icon={icon}
        iconOnly={iconOnly}
        isDisabled={disabled}
        items={items}
      />
    </Tooltip>
  );
}

interface FormatOption {
  id: string;
  label: string;
  icon?: IconSource;
}
function FormatSelector({
  embedded,
  label,
  value,
  options,
  icon,
  disabled,
  className,
  onChange,
}: {
  embedded: boolean;
  label: string;
  value: string;
  options: FormatOption[];
  icon: IconSource;
  disabled: boolean;
  className?: string;
  onChange(value: string): void;
}) {
  if (embedded)
    return (
      <MenuButton
        label={label}
        hint={`${label}: ${options.find((option) => option.id === value)?.label ?? value}`}
        icon={icon}
        iconOnly
        disabled={disabled}
        items={options.map((option) => ({
          ...option,
          icon: option.icon ?? icon,
          isSelected: option.id === value,
          onAction: () => onChange(option.id),
        }))}
      />
    );
  return (
    <SelectField
      tracksChanges={false}
      menu={[]}
      density="compact"
      labelLayout="control-only"
      size="sm"
      className={className}
      label={label}
      value={value}
      isDisabled={disabled}
      options={options}
      onChange={(next) => onChange(String(next))}
    />
  );
}

const PRESENCE_STATUS_LABELS = { active: 'Active', idle: 'Idle', away: 'Away' } as const;

export function EditorToolbar({ editor: e }: { editor: SpreadsheetEditorController }) {
  const menus = editorMenus(e);
  const disabled = !e.canWrite;
  const chrome = e.props.chrome ?? 'full';
  const embedded = chrome === 'embedded';
  if (chrome === 'minimal') return null;
  return (
    <div className="se-editor__chrome" onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
      <div className="se-editor__header">
        {!embedded && (
          <h2 className="se-editor__title">
            {e.props.title ?? e.workbook.name ?? 'Untitled spreadsheet'}
          </h2>
        )}
        {e.props.session && (
          <SaveStatus
            phase={e.savePhase}
            onRetry={embedded ? undefined : () => void e.save()}
            hideSavedAfter={3000}
            labels={{
              dirty: 'Modified.',
              saving: e.saveAutomatic ? 'Autosaving…' : 'Saving…',
              saved: e.saveAutomatic ? 'Autosaved.' : 'Saved.',
            }}
          />
        )}
        {!!e.peers.length && (
          <AvatarGroup
            aria-label="Active collaborators"
            size="xs"
            layout="strip"
            maxVisible={4}
            overflow="menu"
            people={e.peers.map((person) => ({
              id: person.id,
              name: person.name,
              colorIndex: collaboratorColorIndex(person.id),
              description: PRESENCE_STATUS_LABELS[person.status],
            }))}
            onAction={e.followCollaborator}
          />
        )}
        <Button
          size="sm"
          variant="ghost"
          isSelected={e.showContributors}
          icon={editorGlyphs.contributors}
          label="Contributors"
          tooltip="Contributors"
          onPress={() => e.setShowContributors(!e.showContributors)}
        />
        {e.props.features?.modeControl !== false &&
          (embedded ? (
            <MenuButton
              label="Editing mode"
              hint={`Editing mode: ${e.mode === 'author' ? 'Authoring' : e.mode === 'review' ? 'Reviewing' : 'Reading'}`}
              icon={editorGlyphs[e.mode]}
              iconOnly
              disabled={e.locked}
              items={menus.review
                .slice(0, 3)
                .map((item) => ({ ...item, isSelected: item.id === e.mode }))}
            />
          ) : (
            <SelectField
              tracksChanges={false}
              menu={[]}
              density="compact"
              labelLayout="control-only"
              size="sm"
              label="Editing mode"
              value={e.mode}
              isDisabled={e.locked}
              options={[
                {
                  id: 'author',
                  label: 'Authoring',
                  isDisabled: !e.capabilities.edit,
                },
                {
                  id: 'review',
                  label: 'Reviewing',
                  isDisabled: !e.capabilities.suggest || e.props.features?.suggestions === false,
                },
                {
                  id: 'read',
                  label: 'Reading',
                  isDisabled: !e.capabilities.edit && e.capabilities.suggest,
                },
              ]}
              onChange={(value) => e.setMode(value as typeof e.mode)}
            />
          ))}
        {e.props.session && (
          <Button
            size="sm"
            icon={editorGlyphs.save}
            label="Save"
            tooltip={e.savePhase === 'error' ? 'Retry save' : 'Save'}
            shortcut="Ctrl+S"
            isDisabled={!e.canEdit || e.savePhase === 'saving'}
            isPending={e.savePhase === 'saving'}
            onPress={() => e.save()}
          />
        )}
      </div>
      {!embedded && (
        <Menubar aria-label="Spreadsheet menu" className="se-editor__menubar">
          {(['file', 'edit', 'insert', 'format', 'view', 'review'] as const).map((name) => (
            <MenubarMenu
              key={name}
              label={name[0].toUpperCase() + name.slice(1)}
              items={menus[name]}
            />
          ))}
        </Menubar>
      )}
      <ToolbarRow
        aria-label="Spreadsheet formatting"
        className="se-editor__toolbar"
        overflow="scroll"
        justify="start"
      >
        {embedded && (
          <>
            <MenuButton
              label="Spreadsheet actions"
              variant="secondary"
              icon={editorGlyphs.actions}
              iconOnly
              items={(['file', 'edit', 'insert', 'view', 'review'] as const).map((id) => ({
                id,
                label: id[0].toUpperCase() + id.slice(1),
                ...editorCommandMetadata(id),
                items: menus[id],
              }))}
            />
            <ToolbarRowSeparator />
          </>
        )}
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
        <ToolbarRowSeparator />
        <ToolbarRowGroup className="se-editor__tool-group" aria-label="Text formatting">
          <Button
            size="sm"
            icon={editorGlyphs.bold}
            aria-label="Bold selection"
            tooltip="Bold"
            shortcut="Ctrl+B"
            isSelected={e.style['font-weight'] === 'bold'}
            isDisabled={disabled}
            onPress={() => e.attempt(() => e.toggleFormat('font-weight', 'bold'))}
          />
          <Button
            size="sm"
            icon={editorGlyphs.italic}
            aria-label="Italicize selection"
            tooltip="Italic"
            shortcut="Ctrl+I"
            isSelected={e.style['font-style'] === 'italic'}
            isDisabled={disabled}
            onPress={() => e.attempt(() => e.toggleFormat('font-style', 'italic'))}
          />
          <Button
            size="sm"
            icon={editorGlyphs.strike}
            aria-label="Strikethrough selection"
            tooltip="Strikethrough"
            isSelected={e.style['text-decoration'] === 'line-through'}
            isDisabled={disabled}
            onPress={() => e.attempt(() => e.toggleFormat('text-decoration', 'line-through'))}
          />
          <FormatSelector
            embedded={embedded}
            label="Font family"
            className="se-editor__font-family"
            value={e.style['font-family'] ?? 'sans-serif'}
            disabled={disabled}
            icon={editorGlyphs.fontFamily}
            options={['sans-serif', 'serif', 'monospace', 'Arial', 'Georgia', 'Verdana'].map(
              (id) => ({ id, label: id }),
            )}
            onChange={(value) => e.attempt(() => e.format({ 'font-family': value }))}
          />
          <EditorFontSize
            embedded={embedded}
            value={e.style['font-size'] ?? '12px'}
            disabled={disabled}
            onChange={(value) => e.attempt(() => e.format({ 'font-size': value }))}
          />
          <EditorColorControl
            label="Text color"
            icon={editorGlyphs.textColor}
            value={e.style.color ?? null}
            disabled={disabled}
            swatches={['#172b3a', '#b42318', '#175cd3', '#067647']}
            onChange={(value) => e.attempt(() => e.format({ color: value }))}
          />
          <EditorColorControl
            label="Cell fill"
            icon={editorGlyphs.fillColor}
            value={e.style['background-color'] ?? null}
            disabled={disabled}
            swatches={['#ffffff', '#fef0c7', '#d1e9ff', '#d1fadf']}
            onChange={(value) => e.attempt(() => e.format({ 'background-color': value }))}
          />
        </ToolbarRowGroup>
        <ToolbarRowSeparator />
        <ToolbarRowGroup className="se-editor__tool-group" aria-label="Number formatting">
          <FormatSelector
            embedded={embedded}
            label="Number format"
            className="se-editor__number-format"
            value={e.cellFormat.kind}
            disabled={disabled}
            icon={numberFormatGlyphs[e.cellFormat.kind]}
            options={Object.entries(numberFormatGlyphs)
              .filter(
                ([id]) =>
                  !e.props.formatting?.kinds ||
                  e.props.formatting.kinds.includes(id as typeof e.cellFormat.kind),
              )
              .map(([id, icon]) => ({
                id,
                icon,
                label: id[0].toUpperCase() + id.slice(1),
              }))}
            onChange={(value) =>
              e.attempt(() =>
                e.format(
                  {},
                  {
                    kind: value as typeof e.cellFormat.kind,
                    ...(value === 'unit' ? { unit: e.cellFormat.unit ?? 'mm' } : {}),
                  },
                ),
              )
            }
          />
          {e.cellFormat.kind === 'unit' && (
            <FormatSelector
              embedded={embedded}
              label="Display unit"
              className="se-editor__number-format"
              value={e.cellFormat.unit ?? 'mm'}
              disabled={disabled}
              icon={editorGlyphs.unit}
              options={UNIT_PRESETS.map((unit) => ({ id: unit, label: unit }))}
              onChange={(unit) => e.attempt(() => e.format({}, { ...e.cellFormat, unit }))}
            />
          )}
          {e.props.formatting?.decimalPlaces !== false && (
            <>
              <Button
                size="sm"
                icon={editorGlyphs.increaseDecimals}
                label="Increase decimal places"
                tooltip="Increase decimal places"
                isDisabled={disabled}
                onPress={() =>
                  e.attempt(() =>
                    e.format(
                      {},
                      {
                        kind: e.cellFormat.kind === 'general' ? 'number' : e.cellFormat.kind,
                        decimals: Math.min(100, (e.cellFormat.decimals ?? 2) + 1),
                      },
                    ),
                  )
                }
              />
              <Button
                size="sm"
                icon={editorGlyphs.decreaseDecimals}
                label="Decrease decimal places"
                tooltip="Decrease decimal places"
                isDisabled={disabled}
                onPress={() =>
                  e.attempt(() =>
                    e.format(
                      {},
                      {
                        kind: e.cellFormat.kind === 'general' ? 'number' : e.cellFormat.kind,
                        decimals: Math.max(0, (e.cellFormat.decimals ?? 2) - 1),
                      },
                    ),
                  )
                }
              />
            </>
          )}
        </ToolbarRowGroup>
        <ToolbarRowSeparator />
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
            isDisabled={
              !e.canEdit || (e.selection.x1 === e.selection.x2 && e.selection.y1 === e.selection.y2)
            }
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
        {!embedded && (
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
              <Button
                size="sm"
                icon={editorCommandMetadata('context-comment').icon}
                label={`Comments (${e.workbook.comments?.filter((thread) => thread.state === 'open').length ?? 0})`}
                tooltip="Comments"
                onPress={() => e.openPanel('comments')}
                isSelected={e.panelOpen && e.panel === 'comments'}
              />
              <Button
                size="sm"
                icon={editorGlyphs.review}
                label={`Suggestions (${e.workbook.suggestions?.filter((item) => item.status === 'pending').length ?? 0})`}
                tooltip="Suggestions"
                onPress={() => e.openPanel('suggestions')}
                isSelected={e.panelOpen && e.panel === 'suggestions'}
              />
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
        )}
      </ToolbarRow>
    </div>
  );
}
