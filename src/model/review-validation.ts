import { assertNumericValue } from '../calculation/values.js';
import { assertCellFormat } from './cell-format.js';
import type { ReviewActor, SpreadsheetComment, SpreadsheetSuggestion } from './review.js';
import { parseSpreadsheetAddress, parseSpreadsheetRange } from './spreadsheet-addressing.js';
import type { Workbook, Worksheet } from './workbook.js';

/** Validates a structural suggestion's before or after worksheets as a workbook of their own. */
export type SnapshotCheck = (snapshot: { schemaVersion: 1; worksheets: Worksheet[] }) => void;

type Thread = SpreadsheetComment;
type Post = Thread['entries'][number];
type Suggestion = SpreadsheetSuggestion;
type SuggestedCell = Suggestion['changes'][number];

const validId = (value: unknown) => typeof value === 'string' && value.length > 0;
const validActor = (value: ReviewActor) =>
  value && typeof value.id === 'string' && value.id.length > 0 && typeof value.name === 'string';
const validInput = (value: unknown) =>
  value === null ||
  typeof value === 'boolean' ||
  (typeof value === 'string' && value.length <= 10_000) ||
  (typeof value === 'number' && Number.isFinite(value));
const validTime = (value: string) =>
  typeof value === 'string' && Number.isFinite(Date.parse(value));
const unique = (ids: string[]) =>
  ids.every((id) => typeof id === 'string' && id.length > 0) && new Set(ids).size === ids.length;

/** A thread anchors to a range on a sheet, or to the whole workbook with neither. */
function invalidPlace(thread: Thread): boolean {
  if (thread.scope !== undefined && !['range', 'workbook'].includes(thread.scope)) return true;
  return thread.scope === 'workbook'
    ? thread.sheetId !== undefined || thread.range !== undefined
    : typeof thread.sheetId !== 'string' || !parseSpreadsheetRange(thread.range ?? '');
}

function invalidResolution(thread: Thread): boolean {
  const { decision } = thread;
  return (
    (thread.suggestion !== undefined && !validId(thread.suggestion)) ||
    (thread.resolvedBy !== undefined && !validId(thread.resolvedBy)) ||
    (thread.resolvedAt !== undefined && !validTime(thread.resolvedAt)) ||
    (decision !== undefined &&
      (!decision ||
        !['accept', 'reject'].includes(decision.decision) ||
        !validId(decision.suggestion) ||
        !validId(decision.by) ||
        !validTime(decision.at)))
  );
}

function invalidThread(thread: Thread): boolean {
  return (
    !thread ||
    typeof thread.id !== 'string' ||
    invalidPlace(thread) ||
    !['open', 'resolved'].includes(thread.state) ||
    !Array.isArray(thread.entries) ||
    !thread.entries.length ||
    thread.entries.length > 10_000 ||
    !unique(thread.entries.map((post) => post?.id)) ||
    invalidResolution(thread)
  );
}

const invalidPost = (post: Post) =>
  !post ||
  typeof post.id !== 'string' ||
  !validId(post.author) ||
  typeof post.text !== 'string' ||
  post.text.length > 50_000 ||
  !validTime(post.createdAt) ||
  (post.editedAt !== undefined && !validTime(post.editedAt));

const invalidReactions = ({ reactions }: Post) =>
  reactions !== undefined &&
  (!reactions ||
    typeof reactions !== 'object' ||
    Array.isArray(reactions) ||
    Object.entries(reactions).some(
      ([emoji, users]) => !emoji || emoji.length > 128 || !Array.isArray(users) || !unique(users),
    ));

function checkComments(workbook: Workbook) {
  if (workbook.comments !== undefined && !Array.isArray(workbook.comments))
    throw new TypeError('Invalid comments.');
  const comments = workbook.comments ?? [];
  if (comments.length > 10_000 || !unique(comments.map((thread) => thread?.id)))
    throw new TypeError('Invalid comment ids or count.');
  for (const thread of comments) {
    if (invalidThread(thread)) throw new TypeError('Invalid comment thread.');
    if (thread.entries.some(invalidPost)) throw new TypeError('Invalid comment entry.');
    if (thread.entries.some(invalidReactions)) throw new TypeError('Invalid comment reactions.');
  }
}

function invalidSuggestionMetadata(suggestion: Suggestion): boolean {
  return (
    !validActor(suggestion.author) ||
    !validTime(suggestion.createdAt) ||
    (suggestion.updatedAt !== undefined && !validTime(suggestion.updatedAt)) ||
    (suggestion.commentThreadId !== undefined &&
      (typeof suggestion.commentThreadId !== 'string' || !suggestion.commentThreadId))
  );
}

const invalidSuggestion = (suggestion: Suggestion) =>
  !suggestion ||
  typeof suggestion.id !== 'string' ||
  invalidSuggestionMetadata(suggestion) ||
  !['pending', 'accepted', 'rejected'].includes(suggestion.status) ||
  typeof suggestion.summary !== 'string' ||
  !Array.isArray(suggestion.changes) ||
  suggestion.changes.length > 50_000;

const invalidCell = (change: SuggestedCell) =>
  !change ||
  typeof change.sheetId !== 'string' ||
  !parseSpreadsheetAddress(change.address) ||
  !validInput(change.before) ||
  !validInput(change.after) ||
  (change.beforeStyle !== undefined && typeof change.beforeStyle !== 'string') ||
  (change.afterStyle !== undefined && typeof change.afterStyle !== 'string');

/** Canonical inputs must be valid numbers that override no formula. */
function checkInputValues(change: SuggestedCell) {
  for (const field of ['beforeInputValue', 'afterInputValue'] as const) {
    if (change[field] === undefined) continue;
    assertNumericValue(change[field], 10_000);
    const input = field === 'beforeInputValue' ? change.before : change.after;
    if (typeof input === 'string' && input.trimStart().startsWith('='))
      throw new TypeError('A suggested canonical input cannot override a formula.');
  }
}

function checkFormats(change: SuggestedCell) {
  for (const field of ['beforeFormat', 'afterFormat'] as const)
    if (change[field] !== undefined) assertCellFormat(change[field]);
}

function checkStructure(
  structure: NonNullable<Suggestion['structure']>,
  assertSnapshot: SnapshotCheck,
) {
  assertSnapshot({ schemaVersion: 1, worksheets: structure.before });
  assertSnapshot({ schemaVersion: 1, worksheets: structure.after });
  const { operation } = structure;
  if (!operation) return;
  const { change } = operation;
  if (
    typeof operation.sheetId !== 'string' ||
    !change ||
    !['row', 'column'].includes(change.axis) ||
    !['insert', 'delete'].includes(change.kind) ||
    !Number.isSafeInteger(change.index) ||
    change.index < 0 ||
    !Number.isSafeInteger(change.count) ||
    change.count < 1
  )
    throw new TypeError('Invalid suggested structure operation.');
}

function checkSuggestions(workbook: Workbook, assertSnapshot: SnapshotCheck) {
  if (workbook.suggestions !== undefined && !Array.isArray(workbook.suggestions))
    throw new TypeError('Invalid suggestions.');
  const suggestions = workbook.suggestions ?? [];
  if (suggestions.length > 10_000 || !unique(suggestions.map((item) => item?.id)))
    throw new TypeError('Invalid suggestion ids or count.');
  for (const suggestion of suggestions) {
    if (invalidSuggestion(suggestion)) throw new TypeError('Invalid suggestion.');
    if (suggestion.changes.some(invalidCell)) throw new TypeError('Invalid suggested cell.');
    suggestion.changes.forEach(checkInputValues);
    suggestion.changes.forEach(checkFormats);
    if (suggestion.structure) checkStructure(suggestion.structure, assertSnapshot);
  }
}

/** Validates comments and suggestions; structural snapshots go through `assertSnapshot`. */
export function checkReviewData(workbook: Workbook, assertSnapshot: SnapshotCheck): void {
  checkComments(workbook);
  checkSuggestions(workbook, assertSnapshot);
}
