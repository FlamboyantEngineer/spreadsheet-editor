import type { Dispatch, SetStateAction } from 'react';
import { safeCoord } from './jspreadsheet-bridge.js';
import type {
  JSpreadsheetCellValue,
  JSpreadsheetConfig,
  JSpreadsheetSelectionRange,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';

/** The config the engine is mounted with, and the key that forces a fresh mount. */
export interface JSpreadsheetMountState {
  config: JSpreadsheetConfig;
  nonce: number;
}

/** The selection as the engine's `getSelection` tuple reports it, for engines without lists. */
function selectionTuple(
  worksheet: JSpreadsheetWorksheetInstance | null,
): JSpreadsheetSelectionRange | null {
  if (!worksheet || typeof worksheet.getSelection !== 'function') return null;
  const selection = worksheet.getSelection();
  if (!Array.isArray(selection) || selection.length !== 4) return null;
  const [x1, y1, x2, y2] = selection.map((value: unknown) => safeCoord(value));
  return { x1, y1, x2, y2 };
}

export function currentSelection(
  worksheet: JSpreadsheetWorksheetInstance | null,
): JSpreadsheetSelectionRange | null {
  if (
    !worksheet ||
    typeof worksheet.getSelectedColumns !== 'function' ||
    typeof worksheet.getSelectedRows !== 'function'
  )
    return selectionTuple(worksheet);
  const columns = worksheet.getSelectedColumns(true) as number[] | undefined;
  const rows = worksheet.getSelectedRows(true) as number[] | undefined;
  if (!Array.isArray(columns) || !Array.isArray(rows) || columns.length === 0 || rows.length === 0)
    return null;
  return {
    x1: Math.min(...columns),
    x2: Math.max(...columns),
    y1: Math.min(...rows),
    y2: Math.max(...rows),
  };
}

export function readCellValue(
  worksheet: JSpreadsheetWorksheetInstance | null,
  { row, col }: { row: number; col: number },
): JSpreadsheetCellValue {
  if (!worksheet) return null;
  if (typeof worksheet.getValueFromCoords === 'function')
    return worksheet.getValueFromCoords(col, row) as JSpreadsheetCellValue;
  if (typeof worksheet.getValue === 'function')
    return worksheet.getValue(col, row) as JSpreadsheetCellValue;
  return null;
}

type Worksheet = JSpreadsheetWorksheetInstance | null;
type Cell = { row: number; col: number };

export function writeCellValue(
  worksheet: Worksheet,
  { row, col }: Cell,
  value: JSpreadsheetCellValue,
) {
  if (!worksheet) return;
  if (typeof worksheet.setValueFromCoords === 'function') {
    worksheet.setValueFromCoords(col, row, value);
  } else if (typeof worksheet.setValue === 'function') {
    worksheet.setValue([[col, row]], value);
  }
}

/**
 * Replaces the dataset in place. Engines without an in-place setData rebuild the
 * mount config with the new dataset and force a React remount of the vendor
 * <Spreadsheet> (a `key` bump). The mount effect's cleanup tears the old engine
 * down and its re-run adopts the freshly mounted one.
 */
export function replaceData(
  worksheet: Worksheet,
  data: ReadonlyArray<ReadonlyArray<JSpreadsheetCellValue>>,
  setMountState: Dispatch<SetStateAction<JSpreadsheetMountState>>,
) {
  if (worksheet && typeof worksheet.setData === 'function') {
    worksheet.setData(data);
    return;
  }
  setMountState((current) => ({
    config: { ...current.config, data },
    nonce: current.nonce + 1,
  }));
}

export function refreshWorksheet(worksheet: Worksheet) {
  if (typeof worksheet?.refresh === 'function') worksheet.refresh();
}
