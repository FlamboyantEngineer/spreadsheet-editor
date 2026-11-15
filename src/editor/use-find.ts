import { useMemo, useState } from 'react';
import { calculationCellKey } from '../calculation/publication.js';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import type { CalculationResult } from '../calculation/evaluator.js';
import type { DisplayOptions } from '../calculation/format.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { assertWorkbook, copyWorkbook, type Workbook } from '../model/workbook.js';
import type { CellTargets } from './edit-targets.js';
import type { CellPresentationSnapshot } from './public-host.js';

export interface FindMatch {
  sheetId: string;
  x: number;
  y: number;
  authored: boolean;
}

/** Find state, and the cells whose authored input or shown value contains the query. */
export function useFind(
  workbook: Workbook,
  result: CalculationResult | null | undefined,
  readyCells: ReadonlySet<string>,
  presentation: CellPresentationSnapshot | undefined,
  display: DisplayOptions | undefined,
) {
  const [showFind, setShowFind] = useState(false),
    [query, setQueryState] = useState(''),
    [replacement, setReplacement] = useState('');
  const [activeMatch, setActiveMatch] = useState<string | null>(null);
  function setQuery(value: string) {
    setQueryState(value);
    setActiveMatch(null);
  }
  const matches = useMemo(() => {
    if (!query) return [];
    const matches: FindMatch[] = [];
    workbook.worksheets.forEach((sheet) =>
      sheet.data.forEach((row, y) =>
        row.forEach((value, x) => {
          const authored = String(value ?? '')
            .toLowerCase()
            .includes(query.toLowerCase());
          const key = calculationCellKey(sheet.id, x, y),
            ready = readyCells.has(key);
          const rawResult =
            result?.worksheets.find((entry) => entry.id === sheet.id)?.values[y]?.[x] ?? null;
          const cellPresentation = presentation?.cells[key];
          const text = ready
            ? (cellPresentation?.text ??
              formatWorksheetCell(
                rawResult,
                sheet.formats?.[formatSpreadsheetAddress(x, y)] ??
                  (cellPresentation?.defaultUnit
                    ? { kind: 'unit', unit: cellPresentation.defaultUnit }
                    : undefined),
                display,
              ))
            : '';
          if (authored || text.toLowerCase().includes(query.toLowerCase()))
            matches.push({ sheetId: sheet.id, x, y, authored });
        }),
      ),
    );
    return matches;
  }, [workbook, query, result, readyCells, presentation, display]);
  const matchIndex =
    activeMatch === null
      ? -1
      : matches.findIndex(
          (match) => calculationCellKey(match.sheetId, match.x, match.y) === activeMatch,
        );
  return {
    showFind,
    setShowFind,
    query,
    setQuery,
    replacement,
    setReplacement,
    setActiveMatch,
    matches,
    matchIndex,
  };
}

/** Writes `text` over the query in each authored match, all in one copy of `base` (SE-8). */
export function replaceMatches(
  base: Workbook,
  targets: readonly FindMatch[],
  query: string,
  text: string,
  all: boolean,
): Workbook {
  const next = copyWorkbook(base),
    pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), all ? 'gi' : 'i');
  for (const match of targets) {
    const target = next.worksheets.find((sheet) => sheet.id === match.sheetId)!;
    target.data[match.y][match.x] = String(target.data[match.y][match.x] ?? '').replace(
      pattern,
      () => text,
    );
    if (target.inputValues) delete target.inputValues[formatSpreadsheetAddress(match.x, match.y)];
  }
  assertWorkbook(next);
  return next;
}

/** The cells `targets` write, by worksheet id. */
export function matchTargets(targets: readonly FindMatch[]): CellTargets {
  return targets.reduce<CellTargets>((cells, match) => {
    const list = (cells.get(match.sheetId) as string[] | undefined) ?? [];
    list.push(formatSpreadsheetAddress(match.x, match.y));
    return cells.set(match.sheetId, list);
  }, new Map());
}
