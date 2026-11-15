'use client';

import { DropdownMenu } from '@flamboyantengineer/ui-components/dropdown-menu';
import type { IconSource } from '@flamboyantengineer/ui-components/icon';
import type { MenuEntry } from '@flamboyantengineer/ui-components/menu';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { Tooltip } from '@flamboyantengineer/ui-components/tooltip';
import type { editorMenus } from './editor-menus.js';
import type { SpreadsheetEditorController } from './use-editor.js';

/** What every toolbar section reads. */
export interface ToolbarSectionProps {
  e: SpreadsheetEditorController;
  menus: ReturnType<typeof editorMenus>;
  /** Cell writes are unavailable. */
  disabled: boolean;
  embedded: boolean;
}

export function MenuButton({
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
export function FormatSelector({
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
