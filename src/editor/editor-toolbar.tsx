'use client';

import {
  UsersGlyph,
  EditGlyph,
  ReplyGlyph,
  EyeGlyph,
  UndoGlyph,
  RedoGlyph,
  SaveGlyph,
} from '@flamboyantengineer/ui-components/core';
import { Button } from '@flamboyantengineer/ui-components/button';
import {
  Menu,
  MenuPopover,
  MenuTrigger,
  type MenuEntry,
} from '@flamboyantengineer/ui-components/menu';
import { Menubar, MenubarMenu } from '@flamboyantengineer/ui-components/menubar';
import { AvatarGroup } from '@flamboyantengineer/ui-components/avatar-group';
import { collaboratorColorIndex } from '../model/presence.js';
import { ToolbarScroller } from '@flamboyantengineer/ui-components/toolbar-scroller';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { ColorField } from '@flamboyantengineer/ui-components/color-field';
import { SaveStatus } from '@flamboyantengineer/ui-components/save-status';
import type { SpreadsheetEditorController } from './use-editor.js';
import { UNIT_PRESETS } from '../calculation/units.js';
import { BorderGlyph } from './border-glyph.js';
import { editorMenus } from './editor-menus.js';

function MenuButton({
  label,
  items,
  disabled,
  grid,
}: {
  label: string;
  items: MenuEntry[];
  disabled?: boolean;
  grid?: boolean | 'border';
}) {
  return (
    <MenuTrigger>
      <Button size="sm" isDisabled={disabled}>
        {label}
      </Button>
      <MenuPopover>
        <Menu
          className={
            grid === 'border'
              ? 'se-editor__border-grid'
              : grid
                ? 'se-editor__format-grid'
                : undefined
          }
          aria-label={label}
          items={items}
        />
      </MenuPopover>
    </MenuTrigger>
  );
}

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
        {e.props.onSave && (
          <SaveStatus
            phase={e.savePhase}
            onRetry={() => void e.save()}
            hideSavedAfter={3000}
            labels={{
              dirty: 'Modified.',
              saving: e.saveKind === 'autosave' ? 'Autosaving…' : 'Saving snapshot…',
              saved: e.saveKind === 'autosave' ? 'Autosaved.' : 'Snapshot saved.',
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
              description: person.activity ?? 'Viewing',
            }))}
            onAction={e.followCollaborator}
          />
        )}
        <Button
          size="sm"
          variant="ghost"
          isSelected={e.showContributors}
          icon={UsersGlyph}
          label="Contributors"
          tooltip="Contributors"
          onPress={() => e.setShowContributors(!e.showContributors)}
        ></Button>
        {e.props.features?.modeControl !== false &&
          (embedded ? (
            <MenuTrigger>
              <Button
                size="sm"
                variant="ghost"
                icon={e.mode === 'author' ? EditGlyph : e.mode === 'review' ? ReplyGlyph : EyeGlyph}
                label="Editing mode"
                tooltip={`Editing mode: ${e.mode === 'author' ? 'Authoring' : e.mode === 'review' ? 'Reviewing' : 'Reading'}`}
                isDisabled={e.props.readOnly}
              />
              <MenuPopover>
                <Menu aria-label="Editing mode" items={menus.review.slice(0, 3)} />
              </MenuPopover>
            </MenuTrigger>
          ) : (
            <SelectField
              density="compact"
              labelLayout="control-only"
              size="sm"
              label="Editing mode"
              value={e.mode}
              isDisabled={e.props.readOnly}
              options={[
                {
                  id: 'author',
                  label: 'Authoring',
                  isDisabled: e.props.permissions?.edit === false,
                },
                {
                  id: 'review',
                  label: 'Reviewing',
                  isDisabled:
                    e.props.permissions?.suggest === false ||
                    e.props.features?.suggestions === false,
                },
                {
                  id: 'read',
                  label: 'Reading',
                  isDisabled:
                    e.props.permissions?.edit === false && e.props.permissions?.suggest !== false,
                },
              ]}
              onChange={(value) => e.setMode(value as typeof e.mode)}
            />
          ))}
        {e.props.onSave && (
          <Button
            size="sm"
            startIcon={SaveGlyph}
            isDisabled={!e.canEdit || e.savePhase === 'saving'}
            onPress={() => e.save()}
          >
            Save
          </Button>
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
      <ToolbarScroller aria-label="Spreadsheet formatting" className="se-editor__toolbar">
        {embedded && (
          <MenuButton
            label="Spreadsheet actions"
            items={(['file', 'edit', 'insert', 'view', 'review'] as const).map((id) => ({
              id,
              label: id[0].toUpperCase() + id.slice(1),
              items: menus[id],
            }))}
          />
        )}
        <div className="se-editor__tool-group" role="group" aria-label="History">
          <Button
            size="sm"
            icon={UndoGlyph}
            aria-label="Undo"
            variant="ghost"
            onPress={e.undo}
            isDisabled={disabled || !e.history.past.length}
            tooltip="Undo"
          ></Button>
          <Button
            size="sm"
            icon={RedoGlyph}
            aria-label="Redo"
            variant="ghost"
            onPress={e.redo}
            isDisabled={disabled || !e.history.future.length}
            tooltip="Redo"
          ></Button>
        </div>
        <div className="se-editor__tool-group" role="group" aria-label="Text formatting">
          <Button
            size="sm"
            aria-label="Bold selection"
            tooltip="Bold"
            isSelected={e.style['font-weight'] === 'bold'}
            isDisabled={disabled}
            onPress={() => e.attempt(() => e.toggleFormat('font-weight', 'bold'))}
          >
            <strong>B</strong>
          </Button>
          <Button
            size="sm"
            aria-label="Italicize selection"
            tooltip="Italic"
            isSelected={e.style['font-style'] === 'italic'}
            isDisabled={disabled}
            onPress={() => e.attempt(() => e.toggleFormat('font-style', 'italic'))}
          >
            <em>I</em>
          </Button>
          <Button
            size="sm"
            aria-label="Strikethrough selection"
            tooltip="Strikethrough"
            isSelected={e.style['text-decoration'] === 'line-through'}
            isDisabled={disabled}
            onPress={() => e.attempt(() => e.toggleFormat('text-decoration', 'line-through'))}
          >
            <s>S</s>
          </Button>
          <SelectField
            density="compact"
            labelLayout="control-only"
            size="sm"
            className="se-editor__font-family"
            label="Font family"
            value={e.style['font-family'] ?? 'sans-serif'}
            isDisabled={disabled}
            options={['sans-serif', 'serif', 'monospace', 'Arial', 'Georgia', 'Verdana'].map(
              (id) => ({ id, label: id }),
            )}
            onChange={(value) => e.attempt(() => e.format({ 'font-family': String(value) }))}
          />
          <SelectField
            density="compact"
            labelLayout="control-only"
            size="sm"
            className="se-editor__font-size"
            label="Font size"
            value={e.style['font-size'] ?? '12px'}
            isDisabled={disabled}
            options={[10, 11, 12, 14, 16, 18, 24, 32].map((size) => ({
              id: `${size}px`,
              label: String(size),
            }))}
            onChange={(value) => e.attempt(() => e.format({ 'font-size': String(value) }))}
          />
          {!embedded && (
            <>
              <ColorField
                density="compact"
                labelLayout="control-only"
                size="sm"
                label="Text color"
                value={e.style.color ?? null}
                isDisabled={disabled}
                swatches={['#172b3a', '#b42318', '#175cd3', '#067647']}
                onChange={(value) => e.attempt(() => e.format({ color: value }))}
              />
              <ColorField
                density="compact"
                labelLayout="control-only"
                size="sm"
                label="Cell fill"
                value={e.style['background-color'] ?? null}
                isDisabled={disabled}
                swatches={['#ffffff', '#fef0c7', '#d1e9ff', '#d1fadf']}
                onChange={(value) => e.attempt(() => e.format({ 'background-color': value }))}
              />
            </>
          )}
        </div>
        <div className="se-editor__tool-group" role="group" aria-label="Number formatting">
          <SelectField
            density="compact"
            labelLayout="control-only"
            size="sm"
            className="se-editor__number-format"
            label="Number format"
            value={e.cellFormat.kind}
            isDisabled={disabled}
            options={[
              'general',
              'number',
              'percent',
              'currency',
              'scientific',
              'fraction',
              'date',
              'time',
              'datetime',
              'unit',
            ].map((id) => ({ id, label: id[0].toUpperCase() + id.slice(1) }))}
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
            <SelectField
              className="se-editor__number-format"
              label="Display unit"
              density="compact"
              labelLayout="control-only"
              size="sm"
              value={e.cellFormat.unit ?? 'mm'}
              isDisabled={disabled}
              options={UNIT_PRESETS.map((unit) => ({ id: unit, label: unit }))}
              onChange={(unit) =>
                e.attempt(() => e.format({}, { ...e.cellFormat, unit: String(unit) }))
              }
            />
          )}
          <Button
            size="sm"
            aria-label="Increase decimal places"
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
          >
            .00+
          </Button>
          <Button
            size="sm"
            aria-label="Decrease decimal places"
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
          >
            .0−
          </Button>
        </div>
        <div className="se-editor__tool-group" role="group" aria-label="Cell layout">
          <MenuButton label="Align cells" items={menus.align} disabled={disabled} grid />
          <MenuButton
            label="Borders"
            items={menus.borders.slice(0, 6).map((item, index) => ({
              ...item,
              icon: <BorderGlyph side={['top', 'left', 'all', 'right', 'bottom', 'none'][index]} />,
            }))}
            disabled={disabled}
            grid="border"
          />
          <Button
            size="sm"
            isDisabled={disabled}
            isSelected={e.style['white-space'] === 'pre-wrap'}
            onPress={() => e.attempt(() => e.toggleFormat('white-space', 'pre-wrap'))}
          >
            Wrap text
          </Button>
          <Button
            size="sm"
            onPress={() => e.attempt(() => e.merge())}
            isDisabled={
              !e.canEdit || (e.selection.x1 === e.selection.x2 && e.selection.y1 === e.selection.y2)
            }
          >
            Merge cells
          </Button>
          <Button
            size="sm"
            onPress={() => e.attempt(() => e.merge(true))}
            isDisabled={!e.canEdit || !Object.keys(e.sheet.mergeCells ?? {}).length}
          >
            Unmerge cells
          </Button>
        </div>
      </ToolbarScroller>
      {!embedded && (
        <div className="se-editor__actions" role="group" aria-label="Spreadsheet commands">
          <Button size="sm" onPress={() => e.attempt(menus.appendRow)} isDisabled={!e.canEdit}>
            Add row
          </Button>
          <Button size="sm" onPress={() => e.attempt(menus.appendColumn)} isDisabled={!e.canEdit}>
            Add column
          </Button>
          <Button
            size="sm"
            isDisabled={e.props.features?.search === false}
            onPress={() => e.setShowFind(!e.showFind)}
            isSelected={e.showFind}
          >
            Find
          </Button>
          <Button
            size="sm"
            onPress={() => e.openPanel('comments')}
            isSelected={e.panelOpen && e.panel === 'comments'}
          >
            Comments ({e.workbook.comments?.filter((thread) => !thread.resolved).length ?? 0})
          </Button>
          <Button
            size="sm"
            onPress={() => e.openPanel('suggestions')}
            isSelected={e.panelOpen && e.panel === 'suggestions'}
          >
            Suggestions (
            {e.workbook.suggestions?.filter((item) => item.status === 'pending').length ?? 0})
          </Button>
          {e.props.onRefresh && (
            <Button
              isDisabled={!e.canEdit}
              size="sm"
              onPress={() =>
                e.attempt(async () => {
                  await e.refresh();
                })
              }
            >
              Refresh data
            </Button>
          )}
        </div>
      )}
    </div>
  );
}
