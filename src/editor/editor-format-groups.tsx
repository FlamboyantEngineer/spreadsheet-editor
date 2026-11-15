'use client';

import { Button } from '@flamboyantengineer/ui-components/button';
import { ToolbarRowGroup } from '@flamboyantengineer/ui-components/toolbar-row';
import { UNIT_PRESETS } from '../calculation/units.js';
import { editorGlyphs, numberFormatGlyphs } from './editor-commands.js';
import { FormatSelector, type ToolbarSectionProps } from './editor-toolbar-controls.js';
import { EditorColorControl, EditorFontSize } from './text-format-controls.js';

const FONT_FAMILIES = ['sans-serif', 'serif', 'monospace', 'Arial', 'Georgia', 'Verdana'];

/** Bold, italic and strikethrough toggles. */
function TextStyleButtons({ e, disabled }: ToolbarSectionProps) {
  return (
    <>
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
    </>
  );
}

export function TextFormattingGroup(props: ToolbarSectionProps) {
  const { e, disabled, embedded } = props;
  return (
    <ToolbarRowGroup className="se-editor__tool-group" aria-label="Text formatting">
      <TextStyleButtons {...props} />
      <FormatSelector
        embedded={embedded}
        label="Font family"
        className="se-editor__font-family"
        value={e.style['font-family'] ?? 'sans-serif'}
        disabled={disabled}
        icon={editorGlyphs.fontFamily}
        options={FONT_FAMILIES.map((id) => ({ id, label: id }))}
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
  );
}

type FormatKind = ToolbarSectionProps['e']['cellFormat']['kind'];

/** The number format kinds the host allows, with their glyphs. */
const formatOptions = (kinds: readonly FormatKind[] | undefined) =>
  Object.entries(numberFormatGlyphs)
    .filter(([id]) => !kinds || kinds.includes(id as FormatKind))
    .map(([id, icon]) => ({ id, icon, label: id[0].toUpperCase() + id.slice(1) }));

/** Buttons that add or remove one decimal place of the number format. */
function DecimalButtons({ e, disabled }: ToolbarSectionProps) {
  const step = (delta: 1 | -1) =>
    e.attempt(() =>
      e.format(
        {},
        {
          kind: e.cellFormat.kind === 'general' ? 'number' : e.cellFormat.kind,
          decimals:
            delta > 0
              ? Math.min(100, (e.cellFormat.decimals ?? 2) + 1)
              : Math.max(0, (e.cellFormat.decimals ?? 2) - 1),
        },
      ),
    );
  return (
    <>
      <Button
        size="sm"
        icon={editorGlyphs.increaseDecimals}
        label="Increase decimal places"
        tooltip="Increase decimal places"
        isDisabled={disabled}
        onPress={() => step(1)}
      />
      <Button
        size="sm"
        icon={editorGlyphs.decreaseDecimals}
        label="Decrease decimal places"
        tooltip="Decrease decimal places"
        isDisabled={disabled}
        onPress={() => step(-1)}
      />
    </>
  );
}

export function NumberFormattingGroup(props: ToolbarSectionProps) {
  const { e, disabled, embedded } = props;
  return (
    <ToolbarRowGroup className="se-editor__tool-group" aria-label="Number formatting">
      <FormatSelector
        embedded={embedded}
        label="Number format"
        className="se-editor__number-format"
        value={e.cellFormat.kind}
        disabled={disabled}
        icon={numberFormatGlyphs[e.cellFormat.kind]}
        options={formatOptions(e.props.formatting?.kinds)}
        onChange={(value) =>
          e.attempt(() =>
            e.format(
              {},
              {
                kind: value as FormatKind,
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
      {e.props.formatting?.decimalPlaces !== false && <DecimalButtons {...props} />}
    </ToolbarRowGroup>
  );
}
