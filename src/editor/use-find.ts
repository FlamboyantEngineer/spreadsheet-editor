'use client';

import { useMemo, useState } from 'react';
import { formatWorksheetCell } from '../calculation/cell-format.js';
import type { CalculationResult } from '../calculation/evaluator.js';
import type { DisplayOptions } from '../calculation/format.js';
import { calculationCellKey } from '../calculation/publication.js';
import { requireDefined } from '../model/require-defined.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { assertWorkbook, copyWorkbook, type Workbook, type Worksheet } from '../model/workbook.js';
import type { CellTargets } from './edit-targets.js';
import type { CellPresentationSnapshot } from './public-host.js';

export interface FindMatch {
  sheetId: string;
  x: number;
  y: number;
  authored: boolean;
}

/** What the find pane searches besides authored input: the values the grid shows. */
interface FindSources {
  result: CalculationResult | null | undefined;
  readyCells: ReadonlySet<string>;
  presentation: CellPresentationSnapshot | undefined;
  display: DisplayOptions | undefined;
}

type SheetFindSources = Omit<FindSources, 'result'> & {
  values: CalculationResult['worksheets'][number]['values'] | undefined;
};

const noMatches: FindMatch[] = [];

/** The text the grid shows for a cell, or nothing while its value is not ready. */
function shownText(
  sheet: Worksheet,
  { x, y }: { x: number; y: number },
  sources: SheetFindSources,
) {
  const { values, readyCells, presentation, display } = sources;
  const key = calculationCellKey(sheet.id, x, y);
  if (!readyCells.has(key)) return '';
  const cellPresentation = presentation?.cells[key];
  if (cellPresentation?.text !== undefined) return cellPresentation.text;
  const unit = cellPresentation?.defaultUnit;
  return formatWorksheetCell(
    values?.[y]?.[x] ?? null,
    sheet.formats?.[formatSpreadsheetAddress(x, y)] ?? (unit ? { kind: 'unit', unit } : undefined),
    display,
  );
}

function findMatches(
  workbook: Workbook,
  query: string,
  { result, ...sources }: FindSources,
): FindMatch[] {
  const needle = query.toLowerCase();
  const valuesBySheet = new Map(result?.worksheets.map((sheet) => [sheet.id, sheet.values]));
  const matches: FindMatch[] = [];
  for (const sheet of workbook.worksheets) {
    const sheetSources = { ...sources, values: valuesBySheet.get(sheet.id) };
    sheet.data.forEach((row, y) =>
      row.forEach((value, x) => {
        const authored = String(value ?? '')
          .toLowerCase()
          .includes(needle);
        if (authored || shownText(sheet, { x, y }, sheetSources).toLowerCase().includes(needle))
          matches.push({ sheetId: sheet.id, x, y, authored });
      }),
    );
  }
  return matches;
}

/** Find state, and the cells whose authored input or shown value contains the query. */
export function useFind(workbook: Workbook, sources: FindSources, enabled: boolean) {
  const [requested, setShowFind] = useState(false),
    [query, setQueryState] = useState(''),
    [replacement, setReplacement] = useState('');
  const [activeMatch, setActiveMatch] = useState<string | null>(null);
  function setQuery(value: string) {
    setQueryState(value);
    setActiveMatch(null);
  }
  const showFind = requested && enabled;
  const { result, readyCells, presentation, display } = sources;
  const matches = useMemo(
    () =>
      showFind && query
        ? findMatches(workbook, query, { result, readyCells, presentation, display })
        : noMatches,
    [showFind, workbook, query, result, readyCells, presentation, display],
  );
  const matchIndex = useMemo(
    () =>
      activeMatch === null
        ? -1
        : matches.findIndex(
            (match) => calculationCellKey(match.sheetId, match.x, match.y) === activeMatch,
          ),
    [activeMatch, matches],
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

/** What a replace writes: `text` over the first, or every, occurrence of `query`. */
export interface Replacement {
  query: string;
  text: string;
  all: boolean;
}

/** Replaces changed authored text in one copy of `base` (SE-8). */
export function replaceMatches(
  base: Workbook,
  targets: readonly FindMatch[],
  { query, text, all }: Replacement,
): { workbook: Workbook; count: number } | null {
  const pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), all ? 'gi' : 'i');
  const sources = new Map(base.worksheets.map((sheet) => [sheet.id, sheet]));
  const replacements = targets.flatMap(({ sheetId, x, y }) => {
    const sheet = requireDefined(sources.get(sheetId));
    const before = String(sheet.data[y]?.[x] ?? '');
    const value = before.replace(pattern, () => text);
    return value === before ? [] : [{ sheetId, x, y, value }];
  });
  if (!replacements.length) return null;
  const workbook = copyWorkbook(base);
  const destinations = new Map(workbook.worksheets.map((sheet) => [sheet.id, sheet]));
  for (const { sheetId, x, y, value } of replacements) {
    const target = requireDefined(destinations.get(sheetId));
    target.data[y][x] = value;
    if (target.inputValues) delete target.inputValues[formatSpreadsheetAddress(x, y)];
  }
  assertWorkbook(workbook);
  return { workbook, count: replacements.length };
}

/** The cells `targets` write, by worksheet id. */
export function matchTargets(targets: readonly FindMatch[]): CellTargets {
  return targets.reduce<CellTargets>((cells, match) => {
    const list = (cells.get(match.sheetId) as string[] | undefined) ?? [];
    list.push(formatSpreadsheetAddress(match.x, match.y));
    return cells.set(match.sheetId, list);
  }, new Map());
}
