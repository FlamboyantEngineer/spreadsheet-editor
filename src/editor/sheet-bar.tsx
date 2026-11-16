'use client';

import { Button } from '@flamboyantengineer/ui-components/button';
import {
  ArrowLeftGlyph,
  ArrowRightGlyph,
  CopyGlyph,
  DeleteGlyph,
  EditGlyph,
  PlusGlyph,
} from '@flamboyantengineer/ui-components/core';
import type { MenuEntry, MenuItemEntry } from '@flamboyantengineer/ui-components/menu';
import { Tab, TabList } from '@flamboyantengineer/ui-components/tabs';
import type { ReactNode } from 'react';
import type { Worksheet } from '../model/workbook.js';

type WorksheetCommand = 'add' | 'insert' | 'duplicate' | 'delete' | 'left' | 'right';

interface SheetBarProps {
  sheets: readonly Pick<Worksheet, 'id' | 'name'>[];
  selectedId: string;
  canEdit: boolean;
  onRename: (id: string, name: string) => void;
  /** A message refuses the proposed name; null or undefined accepts it. */
  validateName: (id: string, name: string) => string | null | undefined;
  onCommand: (command: WorksheetCommand, id: string) => void;
  scrollButtons?: boolean;
  /** Trailing, fixed chrome such as the grid zoom control. */
  children?: ReactNode;
}

interface WorksheetMenuContext {
  canEdit: boolean;
  sheetCount: number;
  onCommand: (command: WorksheetCommand, id: string) => void;
}

interface WorksheetMenuCommand {
  value: WorksheetCommand;
  label: string;
  icon: MenuItemEntry['icon'];
  isDisabled?: boolean;
}

/** A worksheet tab's menu: its commands around the tab list's standard rename entry. */
function worksheetMenu(
  id: string,
  standard: readonly MenuEntry[],
  { canEdit, sheetCount, onCommand }: WorksheetMenuContext,
): MenuEntry[] {
  const command = ({
    value,
    label,
    icon,
    isDisabled = false,
  }: WorksheetMenuCommand): MenuItemEntry => ({
    id: value,
    label,
    icon,
    isDisabled: !canEdit || isDisabled,
    onAction: () => onCommand(value, id),
  });
  const single = sheetCount < 2;
  const edits: MenuEntry[] = canEdit
    ? standard.map((entry) =>
        entry.id === 'rename' ? { ...entry, shortcut: 'F2', 'aria-label': 'Rename' } : entry,
      )
    : [
        {
          id: 'rename',
          label: 'Rename',
          'aria-label': 'Rename',
          icon: EditGlyph,
          isDisabled: true,
          shortcut: 'F2',
        },
        command({ value: 'duplicate', label: 'Duplicate', icon: CopyGlyph }),
      ];
  return [
    command({ value: 'insert', label: 'Insert worksheet before', icon: PlusGlyph }),
    ...edits,
    { id: 'worksheet-move-separator', type: 'separator' },
    command({ value: 'left', label: 'Move left', icon: ArrowLeftGlyph, isDisabled: single }),
    command({ value: 'right', label: 'Move right', icon: ArrowRightGlyph, isDisabled: single }),
    { id: 'worksheet-delete-separator', type: 'separator' },
    {
      ...command({
        value: 'delete',
        label: 'Delete worksheet',
        icon: DeleteGlyph,
        isDisabled: single,
      }),
      tone: 'danger',
    },
  ];
}

/**
 * The worksheet list within its owner's controlled Tabs root. The same root pairs worksheet ids
 * with TabPanels around the grid; selection and pending grid drafts stay with that owner.
 */
export function SheetBar({
  sheets,
  selectedId,
  canEdit,
  onRename,
  validateName,
  onCommand,
  scrollButtons = true,
  children,
}: SheetBarProps) {
  const menuItems = (key: string | number, standard: readonly MenuEntry[]) =>
    worksheetMenu(String(key), standard, { canEdit, sheetCount: sheets.length, onCommand });
  return (
    <div className="se-editor__sheets">
      <TabList
        className="se-editor__sheet-tabs"
        aria-label="Worksheets"
        overflow="scroll"
        scrollButtons={scrollButtons}
        onRename={canEdit ? (key, name) => onRename(String(key), name) : undefined}
        validateName={(key, name) => validateName(String(key), name)}
        onDuplicate={canEdit ? (key) => onCommand('duplicate', String(key)) : undefined}
        getMenuItems={menuItems}
        renderAfterTab={(key) =>
          String(key) === sheets.at(-1)?.id ? (
            <Button
              size="sm"
              icon={PlusGlyph}
              aria-label="Add worksheet"
              tooltip="Add worksheet"
              isDisabled={!canEdit}
              onPress={() => onCommand('add', selectedId)}
            />
          ) : null
        }
      >
        {sheets.map((sheet) => (
          <Tab key={sheet.id} id={sheet.id} textValue={sheet.name}>
            {sheet.name}
          </Tab>
        ))}
      </TabList>
      {children}
    </div>
  );
}
