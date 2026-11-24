import type {
  JSpreadsheetHostHandle,
  JSpreadsheetWorksheetInstance,
} from './jspreadsheet-types.js';

/** Each grid mounts a single worksheet. */
export function gridWorksheet(
  handle: JSpreadsheetHostHandle | null,
): JSpreadsheetWorksheetInstance | null {
  const instance = handle?.getInstance();
  if (!instance) return null;
  return Array.isArray(instance)
    ? (instance[0] ?? null)
    : 'worksheets' in instance
      ? (instance.worksheets?.[0] ?? null)
      : (instance as JSpreadsheetWorksheetInstance);
}

/** CE's coordinate getter throws when a record is not in the mounted worksheet. */
export function gridCellAt(
  engine: JSpreadsheetWorksheetInstance | null | undefined,
  x: number,
  y: number,
): HTMLTableCellElement | null {
  if (!Number.isSafeInteger(x) || x < 0 || !Number.isSafeInteger(y) || y < 0) return null;
  const cell = engine?.records
    ? engine.records[y]?.[x]?.element
    : engine?.getCellFromCoords?.(x, y);
  return cell?.tagName === 'TD' ? (cell as HTMLTableCellElement) : null;
}
