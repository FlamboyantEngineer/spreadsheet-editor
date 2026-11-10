import { copyWorkbook, type Workbook, type CellProvenance } from './workbook.js';
import { formatSpreadsheetAddress } from './spreadsheet-addressing.js';

export function attributeCellChanges(
  before: Workbook,
  after: Workbook,
  entry: CellProvenance,
): Workbook {
  const next = copyWorkbook(after);
  for (const sheet of next.worksheets) {
    const previous = before.worksheets.find((value) => value.id === sheet.id);
    sheet.data.forEach((row, y) =>
      row.forEach((value, x) => {
        if ((previous?.data[y]?.[x] ?? null) !== (value ?? null))
          (sheet.cellProvenance ??= {})[formatSpreadsheetAddress(x, y)] = structuredClone(entry);
      }),
    );
  }
  return next;
}
