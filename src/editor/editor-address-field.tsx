'use client';

import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { type KeyboardEvent, useState } from 'react';
import { type CellSelection, selectionAddress } from '../model/editing.js';
import { parseSpreadsheetRange } from '../model/spreadsheet-addressing.js';

/** Address parsing must respect the workbook's selectable bounds, including during live input. */
function addressSelection(value: string): CellSelection | null {
  const range = parseSpreadsheetRange(value);
  if (!range || range.end.x > 16_383 || range.end.y > 1_048_575) return null;
  return { x1: range.start.x, y1: range.start.y, x2: range.end.x, y2: range.end.y };
}

type ComposingCheck = (event: KeyboardEvent) => boolean;

/** Stops an Enter or Escape key from reaching the grid. */
export function claimKey(event: KeyboardEvent) {
  event.preventDefault();
  event.stopPropagation();
}

/** The address field: typing a valid address navigates; Enter normalizes it, Escape restores. */
export function AddressField({
  selection,
  isComposing,
  onFocus,
  onNavigate,
}: {
  selection: CellSelection;
  isComposing: ComposingCheck;
  onFocus(): void;
  onNavigate(selection: CellSelection): boolean | void;
}) {
  const [addressDraft, setAddressDraft] = useState<string | null>(null);
  const [addressInvalid, setAddressInvalid] = useState(false);
  const address = selectionAddress(selection);
  const normalizeAddress = () => {
    setAddressDraft(null);
    setAddressInvalid(false);
  };
  const commitAddress = () => {
    const next = addressSelection(addressDraft ?? address);
    if (!next || onNavigate(next) === false) {
      setAddressInvalid(true);
      return;
    }
    setAddressDraft(selectionAddress(next));
    setAddressInvalid(false);
  };
  return (
    <TextField
      className="se-editor__address"
      label="Selected cell or range"
      density="compact"
      labelLayout="control-only"
      tracksChanges={false}
      size="sm"
      value={addressDraft ?? address}
      errorMessage={addressInvalid ? 'Enter a cell or range within this worksheet.' : undefined}
      messageDisplay="tooltip"
      onFocus={onFocus}
      onChange={(value) => {
        setAddressDraft(value);
        const next = addressSelection(value);
        setAddressInvalid(!next || onNavigate(next) === false);
      }}
      onBlur={normalizeAddress}
      onKeyDown={(event) => {
        if (isComposing(event)) return;
        if (event.key === 'Escape') {
          claimKey(event);
          normalizeAddress();
        } else if (event.key === 'Enter') {
          claimKey(event);
          commitAddress();
        }
      }}
    />
  );
}
