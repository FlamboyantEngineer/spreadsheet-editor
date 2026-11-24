import type { WorkbookIdFactory } from './extent.js';
import { annotationOperation, type WorkbookOperation } from './operations.js';
import { requireDefined } from './require-defined.js';
import type { ReviewStamp, SuggestedCellChange } from './review.js';
import { acceptSuggestion, recordProvenance, writeSuggestedCell } from './review-acceptance.js';
import { formatSpreadsheetAddress, parseSpreadsheetAddress } from './spreadsheet-addressing.js';
import type { StructureChange } from './structure.js';
import { copyWorkbook, type Workbook, type Worksheet } from './workbook.js';

type Point = { x: number; y: number };

/** What a review edit proposes: the workbook it would leave, and the line command, if any. */
export interface WorkbookProposal {
  proposed: Workbook;
  summary: string;
  operation?: { sheetId: string; change: StructureChange } | undefined;
}
/** A reviewer's decision on one pending suggestion. */
export interface SuggestionDecision {
  id: string;
  decision: 'accepted' | 'rejected';
  idFactory?: WorkbookIdFactory | undefined;
}

const json = (value: unknown) => JSON.stringify(value);

function structureOf(sheets: Worksheet[]) {
  return sheets.map((sheet) => {
    const shape = { ...sheet };
    delete shape.style;
    delete shape.formats;
    delete shape.cellProvenance;
    delete shape.inputValues;
    delete shape.rowIds;
    delete shape.columnIds;
    delete (shape as Partial<Worksheet>).data;
    return shape;
  });
}

const sameLines = (before: string[] | undefined, after: string[] | undefined) =>
  before === undefined ? after === undefined : before.every((id, index) => after?.[index] === id);

function changedIdentities(previous: Worksheet[], next: Worksheet[]): boolean {
  return previous.some((sheet, index) => {
    const candidate = next[index];
    return (
      !candidate ||
      !sameLines(sheet.rowIds, candidate.rowIds) ||
      !sameLines(sheet.columnIds, candidate.columnIds)
    );
  });
}

/** Every address either version of a sheet authors, metadata first, then values. */
function suggestedAddresses(sheet: Worksheet, after: Worksheet): Set<string> {
  const addresses = new Set([
    ...Object.keys(sheet.style ?? {}),
    ...Object.keys(after.style ?? {}),
    ...Object.keys(sheet.formats ?? {}),
    ...Object.keys(after.formats ?? {}),
    ...Object.keys(sheet.inputValues ?? {}),
    ...Object.keys(after.inputValues ?? {}),
  ]);
  for (const version of [after, sheet])
    version.data.forEach((row, y) =>
      row.forEach((_value, x) => addresses.add(formatSpreadsheetAddress(x, y))),
    );
  return addresses;
}

const cellSide = (sheet: Worksheet, address: string, { x, y }: Point) => ({
  value: sheet.data[y]?.[x] ?? null,
  style: sheet.style?.[address],
  format: sheet.formats?.[address],
  inputValue: sheet.inputValues?.[address],
});

function suggestedChange(
  sheet: Worksheet,
  after: Worksheet,
  address: string,
): SuggestedCellChange | undefined {
  const position = parseSpreadsheetAddress(address);
  if (!position) return undefined;
  const old = cellSide(sheet, address, position),
    next = cellSide(after, address, position);
  if (
    old.value === next.value &&
    old.style === next.style &&
    json(old.format) === json(next.format) &&
    json(old.inputValue) === json(next.inputValue)
  )
    return undefined;
  return {
    sheetId: sheet.id,
    address,
    before: old.value,
    after: next.value,
    beforeStyle: old.style,
    afterStyle: next.style,
    beforeFormat: old.format,
    afterFormat: next.format,
    beforeInputValue: old.inputValue,
    afterInputValue: next.inputValue,
  };
}

/** The cell changes between two workbooks of the same structure. */
const cellChanges = (workbook: Workbook, candidate: Workbook): SuggestedCellChange[] =>
  workbook.worksheets.flatMap((sheet, sheetIndex) => {
    const after = candidate.worksheets[sheetIndex];
    return [...suggestedAddresses(sheet, after)].flatMap(
      (address) => suggestedChange(sheet, after, address) ?? [],
    );
  });

const changeKey = (change: SuggestedCellChange) => JSON.stringify([change.sheetId, change.address]);
const changeKeys = (entries: SuggestedCellChange[]) => entries.map(changeKey).sort().join('|');

/** The actor's own pending cell suggestion on exactly these cells, which a new edit refreshes. */
const pendingMatch = (workbook: Workbook, changes: SuggestedCellChange[], actorId: string) =>
  workbook.suggestions?.find(
    (item) =>
      item.status === 'pending' &&
      !item.outdated &&
      !item.structure &&
      item.author.id === actorId &&
      changeKeys(item.changes) === changeKeys(changes),
  );

/** A refreshed cell keeps its original before state and takes each side the new edit changed. */
function refreshedChange(
  existing: ReadonlyMap<string, SuggestedCellChange>,
  change: SuggestedCellChange,
) {
  const before = requireDefined(existing.get(changeKey(change)));
  return {
    ...before,
    after: change.after !== change.before ? change.after : before.after,
    afterInputValue:
      json(change.afterInputValue) !== json(change.beforeInputValue)
        ? change.afterInputValue
        : before.afterInputValue,
    afterStyle: change.afterStyle !== change.beforeStyle ? change.afterStyle : before.afterStyle,
    afterFormat:
      json(change.afterFormat) !== json(change.beforeFormat)
        ? change.afterFormat
        : before.afterFormat,
  };
}

/** Opens the suggestion's thread over the cells it changes on its first sheet. */
function openSuggestionThread(
  next: Workbook,
  changes: SuggestedCellChange[],
  { summary, stamp }: { summary: string; stamp: ReviewStamp },
) {
  const anchor = changes[0];
  if (!anchor) return;
  const positions = changes
    .filter((change) => change.sheetId === anchor.sheetId)
    .map((change) => requireDefined(parseSpreadsheetAddress(change.address)));
  const x1 = Math.min(...positions.map((p) => p.x)),
    y1 = Math.min(...positions.map((p) => p.y)),
    x2 = Math.max(...positions.map((p) => p.x)),
    y2 = Math.max(...positions.map((p) => p.y));
  const range =
    formatSpreadsheetAddress(x1, y1) +
    (x1 === x2 && y1 === y2 ? '' : ':' + formatSpreadsheetAddress(x2, y2));
  (next.comments ??= []).push({
    id: stamp.id + '-thread',
    sheetId: anchor.sheetId,
    range,
    state: 'open',
    suggestion: stamp.id,
    entries: [
      { id: stamp.id + '-post', author: stamp.actor.id, text: summary, createdAt: stamp.time },
    ],
  });
}

/** A review edit records a proposal; canonical data does not change until acceptance. */
export function proposeWorkbookChange(
  workbook: Workbook,
  { proposed, summary, operation }: WorkbookProposal,
  stamp: ReviewStamp,
): Workbook {
  const next = copyWorkbook(workbook),
    candidate = copyWorkbook(proposed);
  const structural =
    operation !== undefined ||
    json(structureOf(workbook.worksheets)) !== json(structureOf(candidate.worksheets)) ||
    changedIdentities(workbook.worksheets, candidate.worksheets);
  const changes = structural ? [] : cellChanges(workbook, candidate);
  if (!structural && !changes.length) return workbook;
  const existing = structural ? undefined : pendingMatch(next, changes, stamp.actor.id);
  if (existing) {
    const previous = new Map(existing.changes.map((change) => [changeKey(change), change]));
    existing.updatedAt = stamp.time;
    existing.summary = summary;
    existing.changes = changes.map((change) => refreshedChange(previous, change));
    return next;
  }
  next.suggestions = [
    ...(next.suggestions ?? []),
    {
      id: stamp.id,
      author: { ...stamp.actor },
      createdAt: stamp.time,
      status: 'pending',
      commentThreadId: stamp.id + '-thread',
      changes,
      summary,
      ...(structural
        ? {
            structure: {
              before: structuredClone(workbook.worksheets),
              after: candidate.worksheets,
              operation,
            },
          }
        : {}),
    },
  ];
  openSuggestionThread(next, changes, { summary, stamp });
  if (!structural)
    next.worksheets.forEach((sheet, index) => {
      if (sheet.rowIds) sheet.rowIds = candidate.worksheets[index].rowIds;
      if (sheet.columnIds) sheet.columnIds = candidate.worksheets[index].columnIds;
    });
  return next;
}

export function decideSuggestion(
  workbook: Workbook,
  decision: SuggestionDecision,
  stamp: ReviewStamp,
): Workbook {
  return decideSuggestionWithOperations(workbook, decision, stamp).workbook;
}

/** Legacy line proposals retain the operations that move cells and annotation anchors. */
export function decideSuggestionWithOperations(
  workbook: Workbook,
  { id, decision, idFactory }: SuggestionDecision,
  stamp: ReviewStamp,
): { workbook: Workbook; operations?: WorkbookOperation[] } {
  let next = copyWorkbook(workbook),
    suggestion = next.suggestions?.find((item) => item.id === id);
  if (!suggestion || suggestion.status !== 'pending')
    throw new Error('Suggestion is no longer pending.');
  const hasStructure = !!suggestion.structure;
  let operations: WorkbookOperation[] | undefined;
  let movedAnnotations: Workbook | undefined;
  if (decision === 'accepted') {
    const result = acceptSuggestion(next, suggestion, idFactory);
    next = result.workbook;
    operations = result.operations;
    if (operations)
      movedAnnotations = {
        ...next,
        ...structuredClone({ comments: next.comments, suggestions: next.suggestions }),
      };
    suggestion = requireDefined(requireDefined(next.suggestions).find((item) => item.id === id));
    recordProvenance(next, suggestion);
  }
  const threadId = suggestion.commentThreadId;
  const linked = next.comments?.find((thread) => thread.id === threadId);
  if (linked) {
    linked.state = 'resolved';
    linked.resolvedBy = stamp.actor.id;
    linked.resolvedAt = stamp.time;
    linked.decision = {
      decision: decision === 'accepted' ? 'accept' : 'reject',
      suggestion: suggestion.id,
      by: stamp.actor.id,
      at: stamp.time,
    };
  }
  suggestion.status = decision;
  suggestion.decidedBy = { ...stamp.actor };
  suggestion.decidedAt = stamp.time;
  if (operations && movedAnnotations) operations.push(annotationOperation(movedAnnotations, next));
  else if (decision === 'accepted' && hasStructure)
    operations = [
      {
        kind: 'replace',
        reason: 'suggestion',
        before: copyWorkbook(workbook),
        after: copyWorkbook(next),
      },
    ];
  return operations ? { workbook: next, operations } : { workbook: next };
}

function previewCell(next: Workbook, change: SuggestedCellChange) {
  const sheet = next.worksheets.find((sheet) => sheet.id === change.sheetId);
  if (sheet && parseSpreadsheetAddress(change.address)) writeSuggestedCell(sheet, change);
}

/** A reviewer's local display is independent of the canonical workbook and other reviewers. */
export function previewSuggestions(workbook: Workbook, actorId: string): Workbook {
  const own = workbook.suggestions?.filter(
    (item) =>
      item.status === 'pending' && !item.outdated && !item.structure && item.author.id === actorId,
  );
  if (!own?.length) return workbook;
  const next = copyWorkbook(workbook);
  for (const item of own) for (const change of item.changes) previewCell(next, change);
  return next;
}
