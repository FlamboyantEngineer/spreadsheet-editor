import {
  assertWorkbook,
  copyWorkbook,
  DEFAULT_WORKBOOK_LIMITS,
  type Workbook,
  type Worksheet,
} from './workbook.js';
import { worksheetExtent, createWorkbookEntityId } from './extent.js';
import {
  formatSpreadsheetAddress,
  formatSpreadsheetColumnName,
  parseSpreadsheetAddress,
  parseSpreadsheetColumnName,
  parseSpreadsheetRange,
} from './spreadsheet-addressing.js';

export interface StructureChange {
  axis: 'row' | 'column';
  kind: 'insert' | 'delete';
  index: number;
  count: number;
  /** IDs generated once by the command owner for newly inserted lines. */
  ids?: readonly string[];
}

function interval(start: number, end: number, change: StructureChange): [number, number] | null {
  const reversed = start > end,
    left = Math.min(start, end),
    right = Math.max(start, end);
  let a = left,
    b = right;
  if (change.kind === 'insert') {
    if (a >= change.index) a += change.count;
    if (b >= change.index) b += change.count;
  } else {
    const last = change.index + change.count - 1;
    if (a > last) a -= change.count;
    else if (a >= change.index) a = change.index;
    if (b > last) b -= change.count;
    else if (b >= change.index) b = change.index - 1;
    if (a > b) return null;
  }
  return reversed ? [b, a] : [a, b];
}
const REFERENCE =
  /^(?:(?:'((?:[^']|'')*)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d*(?::\$?[A-Za-z]{1,3}\$?[1-9]\d*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9]\d*:\$?[1-9]\d*)(?![A-Za-z0-9_.(])/;

/** Scan references while keeping quoted text, anchors, and all authored numeric digits intact. */
function mapReferences(
  formula: string,
  map: (reference: string, sheet: string | null, prefix: string) => string,
): string {
  if (!formula.trimStart().startsWith('=')) return formula;
  let output = '',
    index = 0;
  while (index < formula.length) {
    if (!/[A-Za-z0-9_.]/.test(formula[index - 1] ?? '')) {
      const match = REFERENCE.exec(formula.slice(index));
      if (match) {
        const sheet = match[1]?.replace(/''/g, "'") ?? match[2] ?? null;
        const prefix = match[0].slice(0, match[0].length - match[3].length);
        output += map(match[3], sheet, prefix);
        index += match[0].length;
        continue;
      }
    }
    const character = formula[index++];
    output += character;
    if (character === '"' || character === "'") {
      while (index < formula.length) {
        const next = formula[index++];
        output += next;
        if (next === character) {
          if (formula[index] === character) output += formula[index++];
          else break;
        }
      }
    }
  }
  return output;
}

function rewriteReference(reference: string, change: StructureChange): string {
  const pieces = reference.split(':'),
    first = pieces[0],
    last = pieces[1] ?? first;
  const cellPattern = /^(\$?)([A-Za-z]+)(\$?)([1-9]\d*)$/;
  const firstCell = cellPattern.exec(first),
    lastCell = cellPattern.exec(last);
  if (firstCell && lastCell) {
    const from =
      change.axis === 'row' ? Number(firstCell[4]) - 1 : parseSpreadsheetColumnName(firstCell[2])!;
    const to =
      change.axis === 'row' ? Number(lastCell[4]) - 1 : parseSpreadsheetColumnName(lastCell[2])!;
    const transformed = interval(from, to, change);
    if (!transformed) return '#REF!';
    const format = (match: RegExpExecArray, value: number) =>
      change.axis === 'row'
        ? `${match[1]}${match[2]}${match[3]}${value + 1}`
        : `${match[1]}${formatSpreadsheetColumnName(value)}${match[3]}${match[4]}`;
    return pieces.length === 1
      ? format(firstCell, transformed[0])
      : `${format(firstCell, transformed[0])}:${format(lastCell, transformed[1])}`;
  }
  const wholeColumns = /^\$?[A-Za-z]+$/.test(first);
  if ((wholeColumns && change.axis !== 'column') || (!wholeColumns && change.axis !== 'row'))
    return reference;
  const from = wholeColumns
    ? parseSpreadsheetColumnName(first.replace('$', ''))!
    : Number(first.replace('$', '')) - 1;
  const to = wholeColumns
    ? parseSpreadsheetColumnName(last.replace('$', ''))!
    : Number(last.replace('$', '')) - 1;
  const transformed = interval(from, to, change);
  if (!transformed) return '#REF!';
  const format = (original: string, value: number) =>
    `${original.startsWith('$') ? '$' : ''}${wholeColumns ? formatSpreadsheetColumnName(value) : value + 1}`;
  return `${format(first, transformed[0])}:${format(last, transformed[1])}`;
}

export function rewriteStructureReferences(
  formula: string,
  currentSheet: string,
  targetSheet: string,
  change: StructureChange,
): string {
  return mapReferences(formula, (reference, sheet, prefix) => {
    if ((sheet ?? currentSheet).toUpperCase() !== targetSheet.toUpperCase())
      return prefix + reference;
    const rewritten = rewriteReference(reference, change);
    return rewritten === '#REF!' ? rewritten : prefix + rewritten;
  });
}

function moveAddress(address: string, change: StructureChange): string | null {
  const cell = parseSpreadsheetAddress(address);
  if (!cell) return null;
  const index = change.axis === 'row' ? cell.y : cell.x,
    next = interval(index, index, change);
  return next
    ? formatSpreadsheetAddress(
        change.axis === 'column' ? next[0] : cell.x,
        change.axis === 'row' ? next[0] : cell.y,
      )
    : null;
}
function moveRange(range: string, change: StructureChange): string | null {
  const parsed = parseSpreadsheetRange(range);
  if (!parsed) return null;
  const start = change.axis === 'row' ? parsed.start.y : parsed.start.x,
    end = change.axis === 'row' ? parsed.end.y : parsed.end.x;
  const moved = interval(start, end, change);
  if (!moved) return null;
  return `${formatSpreadsheetAddress(change.axis === 'column' ? moved[0] : parsed.start.x, change.axis === 'row' ? moved[0] : parsed.start.y)}:${formatSpreadsheetAddress(change.axis === 'column' ? moved[1] : parsed.end.x, change.axis === 'row' ? moved[1] : parsed.end.y)}`;
}

export function changeWorkbookStructure(
  workbook: Workbook,
  sheetId: string,
  change: StructureChange,
): Workbook {
  if (
    !Number.isSafeInteger(change.index) ||
    !Number.isSafeInteger(change.count) ||
    change.index < 0 ||
    change.index > 50_000 ||
    change.count < 1 ||
    change.count > 50_000
  )
    throw new RangeError('Invalid structural change.');
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId);
  if (!sheet) throw new RangeError('Unknown worksheet.');
  const extent = worksheetExtent(sheet, next),
    axisExtent = change.axis === 'row' ? extent.rows : extent.columns;
  if (change.index > axisExtent || (change.kind === 'delete' && change.index >= axisExtent))
    throw new RangeError('Structural change is outside the worksheet extent.');
  if (change.kind === 'delete' && change.index + change.count > axisExtent)
    throw new RangeError('Structural deletion is outside the worksheet extent.');
  change = {
    ...change,
    count: change.kind === 'delete' ? Math.min(change.count, axisExtent - 1) : change.count,
  };
  if (!change.count) return next;
  const columns =
      extent.columns +
      (change.axis === 'column' ? (change.kind === 'insert' ? change.count : -change.count) : 0),
    rows =
      extent.rows +
      (change.axis === 'row' ? (change.kind === 'insert' ? change.count : -change.count) : 0);
  if (columns > 16384 || rows > 1048576 || columns * rows > DEFAULT_WORKBOOK_LIMITS.maxCells)
    throw new RangeError('Worksheet exceeds the cell limit.');
  sheet.minDimensions = [columns, rows];
  let existingIds = change.axis === 'row' ? sheet.rowIds : sheet.columnIds;
  if (!existingIds && change.ids) {
    existingIds = Array.from({ length: axisExtent }, () => createWorkbookEntityId(change.axis));
    if (change.axis === 'row') sheet.rowIds = existingIds;
    else sheet.columnIds = existingIds;
  }
  if (
    change.ids !== undefined &&
    (change.kind !== 'insert' ||
      change.ids.length !== change.count ||
      change.ids.some((id) => typeof id !== 'string' || !id || id.length > 256) ||
      new Set(change.ids).size !== change.ids.length ||
      change.ids.some((id) => existingIds?.includes(id)))
  )
    throw new TypeError('Invalid inserted line identities.');
  if (existingIds)
    existingIds.splice(
      change.index,
      change.kind === 'delete' ? change.count : 0,
      ...(change.kind === 'insert'
        ? (change.ids ??
          Array.from({ length: change.count }, () => createWorkbookEntityId(change.axis)))
        : []),
    );
  if (change.axis === 'row') {
    if (change.kind === 'insert') {
      if (change.index < sheet.data.length)
        sheet.data.splice(change.index, 0, ...Array.from({ length: change.count }, () => []));
    } else sheet.data.splice(change.index, change.count);
    const rowSettings: NonNullable<Worksheet['rows']> = {};
    for (const [index, settings] of Object.entries(sheet.rows ?? {})) {
      const moved = interval(Number(index), Number(index), change);
      if (moved) rowSettings[String(moved[0])] = settings;
    }
    sheet.rows = rowSettings;
  } else {
    sheet.data.forEach((row) => {
      if (change.kind === 'insert') {
        if (change.index < row.length)
          row.splice(change.index, 0, ...Array.from({ length: change.count }, () => null));
      } else row.splice(change.index, change.count);
    });
    if (sheet.columns) {
      if (change.kind === 'insert') {
        if (change.index < sheet.columns.length)
          sheet.columns.splice(
            change.index,
            0,
            ...Array.from({ length: change.count }, () => ({})),
          );
      } else sheet.columns.splice(change.index, change.count);
    }
  }
  if (!sheet.data.length) sheet.data = [['']];
  if (!sheet.data.some((row) => row.length)) sheet.data[0].push(null);
  const remap = <T>(map: Record<string, T> | undefined): Record<string, T> =>
    Object.fromEntries(
      Object.entries(map ?? {}).flatMap(([address, value]) => {
        const next = moveAddress(address, change);
        return next ? [[next, value]] : [];
      }),
    );
  sheet.style = remap(sheet.style);
  sheet.formats = remap(sheet.formats);
  sheet.cellProvenance = remap(sheet.cellProvenance);
  sheet.inputValues = remap(sheet.inputValues);
  const merges: NonNullable<Worksheet['mergeCells']> = {};
  for (const [address, [columns, rows]] of Object.entries(sheet.mergeCells ?? {})) {
    const anchor = parseSpreadsheetAddress(address)!;
    const moved = moveRange(
      `${address}:${formatSpreadsheetAddress(anchor.x + columns - 1, anchor.y + rows - 1)}`,
      change,
    );
    const span = moved && parseSpreadsheetRange(moved);
    if (span)
      merges[formatSpreadsheetAddress(span.start.x, span.start.y)] = [
        span.end.x - span.start.x + 1,
        span.end.y - span.start.y + 1,
      ];
  }
  sheet.mergeCells = merges;
  next.worksheets.forEach((current) =>
    current.data.forEach((row) =>
      row.forEach((value, index) => {
        if (typeof value === 'string')
          row[index] = rewriteStructureReferences(value, current.name, sheet.name, change);
      }),
    ),
  );
  next.comments?.forEach((thread) => {
    if (thread.sheetId === sheetId && !thread.orphaned) {
      const moved = moveRange(thread.range ?? '', change);
      if (moved) thread.range = parseSpreadsheetRange(moved)!.address;
      else thread.orphaned = true;
    }
  });
  next.suggestions?.forEach((item) =>
    item.changes.forEach((cell) => {
      const current = next.worksheets.find((sheet) => sheet.id === cell.sheetId);
      if (cell.sheetId === sheetId) {
        const moved = moveAddress(cell.address, change);
        if (moved) cell.address = moved;
        else item.outdated = true;
      }
      if (current)
        for (const field of ['before', 'after'] as const)
          if (typeof cell[field] === 'string')
            cell[field] = rewriteStructureReferences(
              cell[field] as string,
              current.name,
              sheet.name,
              change,
            );
    }),
  );
  assertWorkbook(next);
  return next;
}

/** Return the normalized name accepted by worksheet renaming, or throw without changing the workbook. */
export function validateWorksheetName(workbook: Workbook, sheetId: string, name: string): string {
  const normalized = name.trim().replace(/\s+/g, ' ');
  if (
    !workbook.worksheets.some((sheet) => sheet.id === sheetId) ||
    !normalized ||
    normalized.length > 256 ||
    /[[\]:*?/\\!]/.test(normalized) ||
    workbook.worksheets.some(
      (other) =>
        other.id !== sheetId &&
        other.name.trim().replace(/\s+/g, ' ').toUpperCase() === normalized.toUpperCase(),
    )
  )
    throw new TypeError('Choose a unique worksheet name without reserved characters.');
  return normalized;
}

export function renameWorksheet(workbook: Workbook, sheetId: string, name: string): Workbook {
  const normalized = validateWorksheetName(workbook, sheetId, name),
    next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId)!;
  const previous = sheet.name;
  sheet.name = normalized;
  next.worksheets.forEach((current) =>
    current.data.forEach((row) =>
      row.forEach((value, index) => {
        if (typeof value === 'string')
          row[index] = mapReferences(value, (reference, target, prefix) =>
            target?.trim().replace(/\s+/g, ' ').toUpperCase() ===
            previous.trim().replace(/\s+/g, ' ').toUpperCase()
              ? `'${normalized.replace(/'/g, "''")}'!${reference}`
              : prefix + reference,
          );
      }),
    ),
  );
  return next;
}

export function deleteWorksheet(workbook: Workbook, sheetId: string): Workbook {
  if (workbook.worksheets.length < 2) throw new Error('Keep at least one worksheet.');
  const next = copyWorkbook(workbook),
    sheet = next.worksheets.find((sheet) => sheet.id === sheetId)!;
  next.worksheets = next.worksheets.filter((sheet) => sheet.id !== sheetId);
  next.worksheets.forEach((current) =>
    current.data.forEach((row) =>
      row.forEach((value, index) => {
        if (typeof value === 'string')
          row[index] = mapReferences(value, (reference, target, prefix) =>
            target?.toUpperCase() === sheet.name.toUpperCase() ? '#REF!' : prefix + reference,
          );
      }),
    ),
  );
  next.comments?.forEach((thread) => {
    if (thread.sheetId === sheetId) thread.orphaned = true;
  });
  next.suggestions?.forEach((item) => {
    if (item.changes.some((cell) => cell.sheetId === sheetId)) item.outdated = true;
  });
  return next;
}

/** Translate a copied formula without touching literals, sheet names or absolute axes. */
export function translateFormulaReferences(formula: string, dx: number, dy: number): string {
  return mapReferences(formula, (reference, _sheet, prefix) => {
    const axis = (token: string, kind: 'column' | 'row'): string | null => {
      const anchored = token.startsWith('$');
      const bare = token.replace('$', '');
      const before = kind === 'column' ? parseSpreadsheetColumnName(bare) : Number(bare) - 1;
      if (before === null) return null;
      const after = before + (anchored ? 0 : kind === 'column' ? dx : dy);
      if (after < 0 || after > (kind === 'column' ? 16383 : 1048575)) return null;
      return (
        (anchored ? '$' : '') +
        (kind === 'column' ? formatSpreadsheetColumnName(after) : String(after + 1))
      );
    };
    const pieces = reference.split(':').map((token) => {
      const cell = /^(\$?[A-Za-z]+)(\$?[1-9]\d*)$/.exec(token);
      if (cell) {
        const x = axis(cell[1], 'column'),
          y = axis(cell[2], 'row');
        return x === null || y === null ? null : x + y;
      }
      return axis(token, /[A-Za-z]/.test(token) ? 'column' : 'row');
    });
    return pieces.some((part) => part === null) ? '#REF!' : prefix + pieces.join(':');
  });
}
