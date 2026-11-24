import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import type { CellProvenance, Workbook, Worksheet } from './workbook.js';

function* changedValues(previous: Worksheet | undefined, sheet: Worksheet): Iterable<string> {
  for (const [y, row] of sheet.data.entries()) {
    const before = previous?.data[y];
    if (row === before) continue;
    for (const [x, value] of row.entries())
      if ((before?.[x] ?? null) !== (value ?? null)) yield formatSpreadsheetAddress(x, y);
  }
}

function changedTargets(
  previous: Worksheet | undefined,
  sheet: Worksheet,
  addresses: Iterable<string>,
) {
  return [...addresses].filter((address) => {
    const cell = parseSpreadsheetAddress(address);
    return (
      cell && (previous?.data[cell.y]?.[cell.x] ?? null) !== (sheet.data[cell.y]?.[cell.x] ?? null)
    );
  });
}

export function attributeCellChanges(
  before: Workbook,
  after: Workbook,
  {
    targets,
    ...entry
  }: CellProvenance & { targets?: ReadonlyMap<string, Iterable<string>> | undefined },
): Workbook {
  let changed = false;
  const worksheets = after.worksheets.map((sheet) => {
    const previous = before.worksheets.find((value) => value.id === sheet.id);
    if (sheet === previous) return sheet;
    const addresses = targets
      ? changedTargets(previous, sheet, targets.get(sheet.id) ?? [])
      : changedValues(previous, sheet);
    let cellProvenance: Worksheet['cellProvenance'];
    for (const address of addresses) {
      cellProvenance ??= { ...sheet.cellProvenance };
      cellProvenance[address] = structuredClone(entry);
    }
    if (!cellProvenance) return sheet;
    changed = true;
    return { ...sheet, cellProvenance };
  });
  return changed ? { ...after, worksheets } : after;
}
