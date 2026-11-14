'use client';

import { useState } from 'react';
import { Button } from '@flamboyantengineer/ui-components/button';
import { ColorField } from '@flamboyantengineer/ui-components/color-field';
import {
  ColorSwatchButton,
  ColorSwatchPicker,
  ColorSwatchPickerItem,
} from '@flamboyantengineer/ui-components/color-swatch';
import type { GlyphComponent } from '@flamboyantengineer/ui-components/core';
import { NumberField } from '@flamboyantengineer/ui-components/number-field';
import { Popover, PopoverTrigger } from '@flamboyantengineer/ui-components/popover';
import { editorGlyphs } from './editor-commands.js';

const fontSizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 60, 72, 96];

function useFormatPopover(disabled: boolean) {
  const [isOpen, onOpenChange] = useState(false);
  if (disabled && isOpen) onOpenChange(false);
  return { isOpen: isOpen && !disabled, onOpenChange };
}

export function EditorFontSize({
  value,
  embedded,
  disabled,
  onChange,
}: {
  value: string;
  embedded: boolean;
  disabled: boolean;
  onChange(value: string): void;
}) {
  const popover = useFormatPopover(disabled);
  const parsed = Number.parseFloat(value);
  const field = (
    <NumberField
      className="se-editor__font-size"
      label="Font size"
      labelLayout="control-only"
      density="compact"
      size="sm"
      tracksChanges={false}
      menu={[]}
      value={Number.isFinite(parsed) ? parsed : 12}
      minValue={1}
      maxValue={400}
      showStepper
      stepValues={fontSizes}
      decrementIcon={editorGlyphs.decreaseFontSize}
      incrementIcon={editorGlyphs.increaseFontSize}
      endAdornment="px"
      isDisabled={disabled}
      onChange={(next) => {
        if (next !== null) onChange(`${next}px`);
      }}
    />
  );
  if (!embedded) return field;
  return (
    <PopoverTrigger {...popover}>
      <Button
        label="Font size"
        icon={editorGlyphs.fontSize}
        size="sm"
        tooltip={`Font size: ${value}`}
        isDisabled={disabled}
      />
      <Popover title="Font size">{field}</Popover>
    </PopoverTrigger>
  );
}

export function EditorColorControl({
  label,
  icon,
  value,
  swatches,
  disabled,
  onChange,
}: {
  label: string;
  icon: GlyphComponent;
  value: string | null;
  swatches: readonly string[];
  disabled: boolean;
  onChange(value: string | null): void;
}) {
  const popover = useFormatPopover(disabled);
  return (
    <PopoverTrigger {...popover}>
      <ColorSwatchButton
        label={label}
        icon={icon}
        color={value}
        tooltip={label}
        size="sm"
        isDisabled={disabled}
      />
      <Popover title={label}>
        {({ close }) => (
          <div className="se-editor__color-palette">
            <ColorSwatchPicker
              aria-label={`${label} palette`}
              value={value ?? ''}
              onChange={(next) => {
                onChange(next);
                close();
              }}
            >
              {swatches.map((color) => (
                <ColorSwatchPickerItem
                  key={color}
                  color={color}
                  colorName={color}
                  isDisabled={disabled}
                />
              ))}
            </ColorSwatchPicker>
            <ColorField
              label={`Custom ${label.toLowerCase()}`}
              value={value}
              onChange={onChange}
              isDisabled={disabled}
              allowsTyping
              formats={['hex', 'rgb', 'hsl']}
              tracksChanges={false}
              menu={[]}
              density="compact"
              size="sm"
            />
            <Button
              label={`Clear ${label.toLowerCase()}`}
              isDisabled={disabled || value === null}
              onPress={() => {
                onChange(null);
                close();
              }}
            />
          </div>
        )}
      </Popover>
    </PopoverTrigger>
  );
}
