'use client';

import { Button } from '@flamboyantengineer/ui-components/button';
import {
  Menu,
  MenuPopover,
  MenuTrigger,
  type MenuEntry,
} from '@flamboyantengineer/ui-components/menu';
import { Menubar, MenubarMenu } from '@flamboyantengineer/ui-components/menubar';
import { ToolbarScroller } from '@flamboyantengineer/ui-components/toolbar-scroller';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { ColorField } from '@flamboyantengineer/ui-components/color-field';
import { SaveStatus } from '@flamboyantengineer/ui-components/save-status';
import type { SpreadsheetEditorController } from './use-editor.js';
import { editorMenus } from './editor-menus.js';

function MenuButton({
  label,
  items,
  disabled,
}: {
  label: string;
  items: MenuEntry[];
  disabled?: boolean;
}) {
  return (
    <MenuTrigger>
      <Button isDisabled={disabled}>{label}</Button>
      <MenuPopover>
        <Menu aria-label={label} items={items} />
      </MenuPopover>
    </MenuTrigger>
  );
}

export function EditorToolbar({ editor: e }: { editor: SpreadsheetEditorController }) {
  const menus = editorMenus(e);
  const disabled = !e.canWrite;
  return (
    <div className="se-editor__chrome" onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
      <div className="se-editor__header">
        <h2 className="se-editor__title">
          {e.props.title ?? e.workbook.name ?? 'Untitled spreadsheet'}
        </h2>
        {e.props.onSave && <SaveStatus phase={e.savePhase} onRetry={e.save} />}
        <SelectField
          label="Editing mode"
          value={e.mode}
          isDisabled={e.props.readOnly}
          options={[
            { id: 'author', label: 'Authoring', isDisabled: e.props.permissions?.edit === false },
            {
              id: 'review',
              label: 'Reviewing',
              isDisabled: e.props.permissions?.suggest === false,
            },
            { id: 'read', label: 'Reading' },
          ]}
          onChange={(value) => e.setMode(value as typeof e.mode)}
        />
        {e.props.onSave && (
          <Button isDisabled={e.props.readOnly || e.savePhase === 'saving'} onPress={e.save}>
            Save
          </Button>
        )}
      </div>
      <Menubar aria-label="Spreadsheet menu" className="se-editor__menubar">
        {(['file', 'edit', 'insert', 'format', 'view', 'review'] as const).map((name) => (
          <MenubarMenu
            key={name}
            label={name[0].toUpperCase() + name.slice(1)}
            items={menus[name]}
          />
        ))}
      </Menubar>
      <ToolbarScroller aria-label="Spreadsheet formatting" className="se-editor__toolbar">
        <Button onPress={e.undo} isDisabled={disabled || !e.history.past.length} tooltip="Undo">
          Undo
        </Button>
        <Button onPress={e.redo} isDisabled={disabled || !e.history.future.length} tooltip="Redo">
          Redo
        </Button>
        <Button
          aria-label="Bold selection"
          tooltip="Bold"
          isSelected={e.style['font-weight'] === 'bold'}
          isDisabled={disabled}
          onPress={() => e.attempt(() => e.toggleFormat('font-weight', 'bold'))}
        >
          <strong>B</strong>
        </Button>
        <Button
          aria-label="Italicize selection"
          tooltip="Italic"
          isSelected={e.style['font-style'] === 'italic'}
          isDisabled={disabled}
          onPress={() => e.attempt(() => e.toggleFormat('font-style', 'italic'))}
        >
          <em>I</em>
        </Button>
        <Button
          aria-label="Strikethrough selection"
          tooltip="Strikethrough"
          isSelected={e.style['text-decoration'] === 'line-through'}
          isDisabled={disabled}
          onPress={() => e.attempt(() => e.toggleFormat('text-decoration', 'line-through'))}
        >
          <s>S</s>
        </Button>
        <SelectField
          className="se-editor__compact-field"
          label="Font family"
          value={e.style['font-family'] ?? 'sans-serif'}
          isDisabled={disabled}
          options={['sans-serif', 'serif', 'monospace', 'Arial', 'Georgia', 'Verdana'].map(
            (id) => ({ id, label: id }),
          )}
          onChange={(value) => e.attempt(() => e.format({ 'font-family': String(value) }))}
        />
        <SelectField
          className="se-editor__compact-field"
          label="Font size"
          value={e.style['font-size'] ?? '14px'}
          isDisabled={disabled}
          options={[10, 11, 12, 14, 16, 18, 24, 32].map((size) => ({
            id: `${size}px`,
            label: String(size),
          }))}
          onChange={(value) => e.attempt(() => e.format({ 'font-size': String(value) }))}
        />
        <ColorField
          label="Text color"
          value={e.style.color ?? null}
          isDisabled={disabled}
          swatches={['#172b3a', '#b42318', '#175cd3', '#067647']}
          onChange={(value) => e.attempt(() => e.format({ color: value }))}
        />
        <ColorField
          label="Cell fill"
          value={e.style['background-color'] ?? null}
          isDisabled={disabled}
          swatches={['#ffffff', '#fef0c7', '#d1e9ff', '#d1fadf']}
          onChange={(value) => e.attempt(() => e.format({ 'background-color': value }))}
        />
        <SelectField
          className="se-editor__compact-field"
          label="Number format"
          value={e.cellFormat.kind}
          isDisabled={disabled}
          options={['general', 'number', 'percent', 'currency', 'scientific', 'fraction'].map(
            (id) => ({ id, label: id[0].toUpperCase() + id.slice(1) }),
          )}
          onChange={(value) =>
            e.attempt(() => e.format({}, { kind: value as typeof e.cellFormat.kind }))
          }
        />
        <Button
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
        <MenuButton label="Align cells" items={menus.align} disabled={disabled} />
        <MenuButton label="Borders" items={menus.borders} disabled={disabled} />
        <Button
          isDisabled={disabled}
          isSelected={e.style['white-space'] === 'pre-wrap'}
          onPress={() => e.attempt(() => e.toggleFormat('white-space', 'pre-wrap'))}
        >
          Wrap text
        </Button>
        <Button
          onPress={() => e.attempt(() => e.merge())}
          isDisabled={
            disabled || (e.selection.x1 === e.selection.x2 && e.selection.y1 === e.selection.y2)
          }
        >
          Merge cells
        </Button>
        <Button
          onPress={() => e.attempt(() => e.merge(true))}
          isDisabled={disabled || !Object.keys(e.sheet.mergeCells ?? {}).length}
        >
          Unmerge cells
        </Button>
      </ToolbarScroller>
      <div className="se-editor__actions" role="group" aria-label="Spreadsheet commands">
        <Button onPress={() => e.attempt(() => e.sheetCommand('add'))} isDisabled={disabled}>
          Add worksheet
        </Button>
        <Button onPress={() => e.attempt(menus.appendRow)} isDisabled={disabled}>
          Add row
        </Button>
        <Button onPress={() => e.attempt(menus.appendColumn)} isDisabled={disabled}>
          Add column
        </Button>
        <Button onPress={() => e.setShowFind(!e.showFind)} isSelected={e.showFind}>
          Find
        </Button>
        <Button
          onPress={() => e.openPanel('comments')}
          isSelected={e.panelOpen && e.panel === 'comments'}
        >
          Comments ({e.workbook.comments?.filter((thread) => !thread.resolved).length ?? 0})
        </Button>
        <Button
          onPress={() => e.openPanel('suggestions')}
          isSelected={e.panelOpen && e.panel === 'suggestions'}
        >
          Suggestions (
          {e.workbook.suggestions?.filter((item) => item.status === 'pending').length ?? 0})
        </Button>
        {e.props.onRefresh && (
          <Button
            onPress={() =>
              e.attempt(async () => {
                await e.props.onRefresh?.();
              })
            }
          >
            Refresh data
          </Button>
        )}
      </div>
    </div>
  );
}
