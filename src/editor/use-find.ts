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
export interface FindSources {
  result: CalculationResult | null | undefined;
  readyCells: ReadonlySet<string>;
  presentation: CellPresentationSnapshot | undefined;
  display: DisplayOptions | undefined;
}

/** The text the grid shows for a cell, or nothing while its value is not ready. */
function shownText(sheet: Worksheet, { x, y }: { x: number; y: number }, sources: FindSources) {
  const { result, readyCells, presentation, display } = sources;
  const key = calculationCellKey(sheet.id, x, y);
  if (!readyCells.has(key)) return '';
  const rawResult =
    result?.worksheets.find((entry) => entry.id === sheet.id)?.values[y]?.[x] ?? null;
  const cellPresentation = presentation?.cells[key];
  if (cellPresentation?.text !== undefined) return cellPresentation.text;
  const unit = cellPresentation?.defaultUnit;
  return formatWorksheetCell(
    rawResult,
    sheet.formats?.[formatSpreadsheetAddress(x, y)] ?? (unit ? { kind: 'unit', unit } : undefined),
    display,
  );
}

function findMatches(workbook: Workbook, query: string, sources: FindSources): FindMatch[] {
  if (!query) return [];
  const needle = query.toLowerCase();
  const matches: FindMatch[] = [];
  for (const sheet of workbook.worksheets)
    sheet.data.forEach((row, y) =>
      row.forEach((value, x) => {
        const authored = String(value ?? '')
          .toLowerCase()
          .includes(needle);
        if (authored || shownText(sheet, { x, y }, sources).toLowerCase().includes(needle))
          matches.push({ sheetId: sheet.id, x, y, authored });
      }),
    );
  return matches;
}

/** Find state, and the cells whose authored input or shown value contains the query. */
export function useFind(workbook: Workbook, sources: FindSources) {
  const [showFind, setShowFind] = useState(false),
    [query, setQueryState] = useState(''),
    [replacement, setReplacement] = useState('');
  const [activeMatch, setActiveMatch] = useState<string | null>(null);
  function setQuery(value: string) {
    setQueryState(value);
    setActiveMatch(null);
  }
  const { result, readyCells, presentation, display } = sources;
  const matches = useMemo(
    () => findMatches(workbook, query, { result, readyCells, presentation, display }),
    [workbook, query, result, readyCells, presentation, display],
  );
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

/** What a replace writes: `text` over the first, or every, occurrence of `query`. */
export interface Replacement {
  query: string;
  text: string;
  all: boolean;
}

/** Writes `text` over the query in each authored match, all in one copy of `base` (SE-8). */
export function replaceMatches(
  base: Workbook,
  targets: readonly FindMatch[],
  { query, text, all }: Replacement,
): Workbook {
  const next = copyWorkbook(base),
    pattern = new RegExp(query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), all ? 'gi' : 'i');
  for (const match of targets) {
    const target = requireDefined(next.worksheets.find((sheet) => sheet.id === match.sheetId));
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
