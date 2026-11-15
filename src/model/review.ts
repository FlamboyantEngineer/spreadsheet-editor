import { assertCellFormat } from './cell-format.js';
import type { NumericValue } from '../calculation/values.js';
import { assertNumericValue } from '../calculation/values.js';
import {
  assertWorkbook,
  copyWorkbook,
  type CellInput,
  type Workbook,
  type Worksheet,
} from './workbook.js';
import {
  formatSpreadsheetAddress,
  parseSpreadsheetAddress,
  parseSpreadsheetRange,
} from './spreadsheet-addressing.js';
import { changeWorkbookStructure, type StructureChange } from './structure.js';

export interface ReviewActor {
  id: string;
  name: string;
}
/** One comment of a thread, shaped like the document editor's comment entries. */
export interface CommentEntry {
  id: string;
  /** The id of the person who wrote it; the host's identity adapter gives their name. */
  author: string;
  text: string;
  /** When it was posted: an ISO 8601 date and time. */
  createdAt: string;
  /** When its text was last changed, if it was. */
  editedAt?: string;
  /** The people who reacted to it, by the emoji they reacted with. */
  reactions?: Record<string, string[]>;
}
/** How the suggestion a thread discusses was decided, by whom and when. */
export interface CommentDecision {
  decision: 'accept' | 'reject';
  suggestion: string;
  /** The id of the person who decided it. */
  by: string;
  at: string;
}
/** A comment thread on a range of cells or on the whole workbook. */
export interface SpreadsheetComment {
  id: string;
  state: 'open' | 'resolved';
  /** The opening comment, then the replies, oldest first. */
  entries: CommentEntry[];
  sheetId?: string;
  range?: string;
  scope?: 'range' | 'workbook';
  /** The suggestion the thread discusses, when it does. */
  suggestion?: string;
  /** How its suggestion was decided, once it was. */
  decision?: CommentDecision;
  /** Who resolved it and when, while it is resolved. */
  resolvedBy?: string;
  resolvedAt?: string;
  orphaned?: boolean;
}
export interface SuggestedCellChange {
  sheetId: string;
  address: string;
  before: CellInput;
  after: CellInput;
  beforeStyle?: string;
  afterStyle?: string;
  beforeInputValue?: NumericValue;
  afterInputValue?: NumericValue;
  beforeFormat?: CellFormat;
  afterFormat?: CellFormat;
}
export interface SpreadsheetSuggestion {
  id: string;
  author: ReviewActor;
  createdAt: string;
  updatedAt?: string;
  commentThreadId?: string;
  status: 'pending' | 'accepted' | 'rejected';
  changes: SuggestedCellChange[];
  structure?: {
    before: Worksheet[];
    after: Worksheet[];
    operation?: { sheetId: string; change: StructureChange };
  };
  summary: string;
  outdated?: boolean;
  decidedBy?: ReviewActor;
  decidedAt?: string;
}
export interface CellFormat {
  kind:
    | 'general'
    | 'number'
    | 'percent'
    | 'currency'
    | 'scientific'
    | 'fraction'
    | 'date'
    | 'time'
    | 'datetime'
    | 'unit';
  locale?: string;
  currency?: string;
  decimals?: number;
  unit?: string;
  inputUnit?: string;
  expectedDimension?: string;
  quantityKind?: string;
}
export interface ReviewStamp {
  id: string;
  time: string;
  actor: ReviewActor;
}

export function addComment(
  workbook: Workbook,
  sheetId: string,
  range: string,
  text: string,
  stamp: ReviewStamp,
): Workbook {
  if (!workbook.worksheets.some((sheet) => sheet.id === sheetId) || !parseSpreadsheetRange(range))
    throw new RangeError('Invalid comment anchor.');
  if (findOpenComment(workbook, sheetId, range))
    throw new Error('An open comment already covers this selection.');
  const next = copyWorkbook(workbook);
  if (!text.trim()) throw new TypeError('A comment needs text.');
  next.comments = [
    ...(next.comments ?? []),
    {
      id: stamp.id,
      sheetId,
      range: parseSpreadsheetRange(range)!.address,
      state: 'open',
      entries: [{ id: stamp.id, author: stamp.actor.id, text: text.trim(), createdAt: stamp.time }],
    },
  ];
  return next;
}

export function updateComment(
  workbook: Workbook,
  id: string,
  update: (thread: SpreadsheetComment) => void,
): Workbook {
  const next = copyWorkbook(workbook),
    thread = next.comments?.find((thread) => thread.id === id);
  if (!thread) throw new RangeError('Unknown comment.');
  update(thread);
  return next;
}

/** Resolve or reopen a thread, recording who resolved it and when (SE-REVIEW-1.7). */
export function setCommentState(
  workbook: Workbook,
  id: string,
  state: SpreadsheetComment['state'],
  stamp: Pick<ReviewStamp, 'time' | 'actor'>,
): Workbook {
  return updateComment(workbook, id, (thread) => {
    thread.state = state;
    if (state === 'resolved') {
      thread.resolvedBy = stamp.actor.id;
      thread.resolvedAt = stamp.time;
    } else {
      delete thread.resolvedBy;
      delete thread.resolvedAt;
    }
  });
}

export function replyToComment(
  workbook: Workbook,
  id: string,
  text: string,
  stamp: ReviewStamp,
): Workbook {
  if (!text.trim()) throw new TypeError('A reply needs text.');
  return updateComment(workbook, id, (thread) => {
    if (thread.state === 'resolved') throw new Error('Reopen the thread before replying.');
    thread.entries.push({
      id: stamp.id,
      author: stamp.actor.id,
      text: text.trim(),
      createdAt: stamp.time,
    });
  });
}

export function editCommentEntry(
  workbook: Workbook,
  threadId: string,
  postId: string,
  text: string,
  actorId: string,
  time: string,
): Workbook {
  if (!text.trim()) throw new TypeError('A comment needs text.');
  return updateComment(workbook, threadId, (thread) => {
    const post = thread.entries.find((post) => post.id === postId);
    if (!post || post.author !== actorId) throw new Error('Only the author can edit this comment.');
    post.text = text.trim();
    post.editedAt = time;
  });
}

export function deleteCommentEntry(
  workbook: Workbook,
  threadId: string,
  postId: string,
  actorId: string,
): Workbook {
  const next = updateComment(workbook, threadId, (thread) => {
    const post = thread.entries.find((post) => post.id === postId);
    if (!post || post.author !== actorId)
      throw new Error('Only the author can delete this comment.');
    thread.entries = thread.entries.filter((post) => post.id !== postId);
  });
  next.comments = next.comments?.filter((thread) => thread.entries.length);
  return next;
}

export function toggleCommentReaction(
  workbook: Workbook,
  threadId: string,
  postId: string,
  emoji: string,
  actorId: string,
): Workbook {
  return updateComment(workbook, threadId, (thread) => {
    const post = thread.entries.find((post) => post.id === postId);
    if (!post) throw new RangeError('Unknown comment entry.');
    const users = post.reactions?.[emoji] ?? [];
    post.reactions = {
      ...post.reactions,
      [emoji]: users.includes(actorId)
        ? users.filter((user) => user !== actorId)
        : [...users, actorId],
    };
  });
}

function structureOf(sheets: Worksheet[]) {
  return sheets.map((sheet) => {
    const shape = { ...sheet };
    delete shape.style;
    delete shape.formats;
    delete shape.cellProvenance;
    delete shape.inputValues;
    delete (shape as Partial<Worksheet>).data;
    return shape;
  });
}

/** A review edit records a proposal; canonical data does not change until acceptance. */
export function proposeWorkbookChange(
  workbook: Workbook,
  proposed: Workbook,
  summary: string,
  stamp: ReviewStamp,
  operation?: { sheetId: string; change: StructureChange },
): Workbook {
  const next = copyWorkbook(workbook),
    candidate = copyWorkbook(proposed);
  const structural =
    operation !== undefined ||
    JSON.stringify(structureOf(workbook.worksheets)) !==
      JSON.stringify(structureOf(candidate.worksheets));
  const changes: SuggestedCellChange[] = [];
  if (!structural)
    workbook.worksheets.forEach((sheet, sheetIndex) => {
      const after = candidate.worksheets[sheetIndex];
      const addresses = new Set([
        ...Object.keys(sheet.style ?? {}),
        ...Object.keys(after.style ?? {}),
        ...Object.keys(sheet.formats ?? {}),
        ...Object.keys(after.formats ?? {}),
        ...Object.keys(sheet.inputValues ?? {}),
        ...Object.keys(after.inputValues ?? {}),
      ]);
      after.data.forEach((row, y) =>
        row.forEach((_value, x) => addresses.add(formatSpreadsheetAddress(x, y))),
      );
      sheet.data.forEach((row, y) =>
        row.forEach((_before, x) => addresses.add(formatSpreadsheetAddress(x, y))),
      );
      for (const address of addresses) {
        const position = parseSpreadsheetAddress(address);
        if (!position) continue;
        const { x, y } = position,
          before = sheet.data[y]?.[x] ?? null,
          value = after.data[y]?.[x] ?? null;
        const beforeStyle = sheet.style?.[address],
          afterStyle = after.style?.[address];
        const beforeFormat = sheet.formats?.[address],
          afterFormat = after.formats?.[address];
        if (
          before !== value ||
          beforeStyle !== afterStyle ||
          JSON.stringify(beforeFormat) !== JSON.stringify(afterFormat) ||
          JSON.stringify(sheet.inputValues?.[address]) !==
            JSON.stringify(after.inputValues?.[address])
        )
          changes.push({
            sheetId: sheet.id,
            address,
            before,
            after: value,
            beforeStyle,
            afterStyle,
            beforeFormat,
            afterFormat,
            beforeInputValue: sheet.inputValues?.[address],
            afterInputValue: after.inputValues?.[address],
          });
      }
    });
  if (!structural && !changes.length) return workbook;
  const keys = (entries: SuggestedCellChange[]) =>
    entries
      .map((change) => `${change.sheetId}:${change.address}`)
      .sort()
      .join('|');
  const existing = !structural
    ? next.suggestions?.find(
        (item) =>
          item.status === 'pending' &&
          !item.outdated &&
          !item.structure &&
          item.author.id === stamp.actor.id &&
          keys(item.changes) === keys(changes),
      )
    : undefined;
  if (existing) {
    existing.updatedAt = stamp.time;
    existing.summary = summary;
    existing.changes = changes.map((change) => {
      const before = existing.changes.find(
        (old) => old.sheetId === change.sheetId && old.address === change.address,
      )!;
      return {
        ...before,
        after: change.after !== change.before ? change.after : before.after,
        afterInputValue:
          JSON.stringify(change.afterInputValue) !== JSON.stringify(change.beforeInputValue)
            ? change.afterInputValue
            : before.afterInputValue,
        afterStyle:
          change.afterStyle !== change.beforeStyle ? change.afterStyle : before.afterStyle,
        afterFormat:
          JSON.stringify(change.afterFormat) !== JSON.stringify(change.beforeFormat)
            ? change.afterFormat
            : before.afterFormat,
      };
    });
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
  const anchor = changes[0];
  if (anchor) {
    const positions = changes
      .filter((change) => change.sheetId === anchor.sheetId)
      .map((change) => parseSpreadsheetAddress(change.address)!);
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
        {
          id: stamp.id + '-post',
          author: stamp.actor.id,
          text: summary,
          createdAt: stamp.time,
        },
      ],
    });
  }
  return next;
}

export function decideSuggestion(
  workbook: Workbook,
  id: string,
  decision: 'accepted' | 'rejected',
  stamp: ReviewStamp,
): Workbook {
  let next = copyWorkbook(workbook),
    suggestion = next.suggestions?.find((item) => item.id === id);
  if (!suggestion || suggestion.status !== 'pending')
    throw new Error('Suggestion is no longer pending.');
  if (decision === 'accepted') {
    if (suggestion.outdated)
      throw new Error('The suggested cells were deleted. Reject this outdated suggestion.');
    if (suggestion.structure) {
      if (JSON.stringify(next.worksheets) !== JSON.stringify(suggestion.structure.before))
        throw new Error('The workbook changed since this structural suggestion.');
      if (suggestion.structure.operation) {
        const { sheetId, change } = suggestion.structure.operation;
        next = changeWorkbookStructure(next, sheetId, change);
        suggestion = next.suggestions!.find((item) => item.id === id)!;
      } else {
        next.worksheets = structuredClone(suggestion.structure.after);
        const ids = new Set(next.worksheets.map((sheet) => sheet.id));
        next.comments?.forEach((thread) => {
          if (thread.scope !== 'workbook' && !ids.has(thread.sheetId ?? '')) thread.orphaned = true;
        });
        next.suggestions?.forEach((item) => {
          if (item.changes.some((cell) => !ids.has(cell.sheetId))) item.outdated = true;
        });
      }
    } else {
      for (const change of suggestion.changes) {
        const sheet = next.worksheets.find((sheet) => sheet.id === change.sheetId),
          cell = parseSpreadsheetAddress(change.address);
        if (
          !sheet ||
          !cell ||
          (sheet.data[cell.y]?.[cell.x] ?? null) !== change.before ||
          sheet.style?.[change.address] !== change.beforeStyle ||
          JSON.stringify(sheet.formats?.[change.address]) !== JSON.stringify(change.beforeFormat) ||
          JSON.stringify(sheet.inputValues?.[change.address]) !==
            JSON.stringify(change.beforeInputValue)
        )
          throw new Error(
            `Cell ${change.address} changed since this suggestion. It has not been overwritten.`,
          );
      }
      for (const change of suggestion.changes) {
        const sheet = next.worksheets.find((sheet) => sheet.id === change.sheetId)!,
          cell = parseSpreadsheetAddress(change.address)!;
        while (sheet.data.length <= cell.y) sheet.data.push([]);
        while (sheet.data[cell.y].length <= cell.x) sheet.data[cell.y].push(null);
        sheet.data[cell.y][cell.x] = change.after;
        sheet.inputValues = { ...sheet.inputValues };
        if (change.afterInputValue === undefined) delete sheet.inputValues[change.address];
        else sheet.inputValues[change.address] = structuredClone(change.afterInputValue);
        sheet.style = { ...sheet.style };
        sheet.formats = { ...sheet.formats };
        if (change.afterStyle === undefined) delete sheet.style[change.address];
        else sheet.style[change.address] = change.afterStyle;
        if (change.afterFormat === undefined) delete sheet.formats[change.address];
        else sheet.formats[change.address] = change.afterFormat;
      }
    }
  }
  if (decision === 'accepted')
    for (const change of suggestion.changes) {
      if (
        change.before === change.after &&
        JSON.stringify(change.beforeInputValue) === JSON.stringify(change.afterInputValue)
      )
        continue;
      const sheet = next.worksheets.find((sheet) => sheet.id === change.sheetId);
      if (sheet)
        (sheet.cellProvenance ??= {})[change.address] = {
          actor: { ...suggestion.author },
          source: 'human',
          at: suggestion.updatedAt ?? suggestion.createdAt,
        };
    }
  const linked = next.comments?.find((thread) => thread.id === suggestion!.commentThreadId);
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
  return next;
}

export function assertReviewData(workbook: Workbook): void {
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
  if (workbook.comments !== undefined && !Array.isArray(workbook.comments))
    throw new TypeError('Invalid comments.');
  if (
    (workbook.comments?.length ?? 0) > 10_000 ||
    !unique((workbook.comments ?? []).map((thread) => thread?.id))
  )
    throw new TypeError('Invalid comment ids or count.');
  for (const thread of workbook.comments ?? []) {
    if (
      !thread ||
      typeof thread.id !== 'string' ||
      (thread.scope !== undefined && !['range', 'workbook'].includes(thread.scope)) ||
      (thread.scope === 'workbook'
        ? thread.sheetId !== undefined || thread.range !== undefined
        : typeof thread.sheetId !== 'string' || !parseSpreadsheetRange(thread.range ?? '')) ||
      !['open', 'resolved'].includes(thread.state) ||
      !Array.isArray(thread.entries) ||
      !thread.entries.length ||
      thread.entries.length > 10_000 ||
      !unique(thread.entries.map((post) => post?.id)) ||
      (thread.suggestion !== undefined && !validId(thread.suggestion)) ||
      (thread.resolvedBy !== undefined && !validId(thread.resolvedBy)) ||
      (thread.resolvedAt !== undefined && !validTime(thread.resolvedAt)) ||
      (thread.decision !== undefined &&
        (!thread.decision ||
          !['accept', 'reject'].includes(thread.decision.decision) ||
          !validId(thread.decision.suggestion) ||
          !validId(thread.decision.by) ||
          !validTime(thread.decision.at)))
    )
      throw new TypeError('Invalid comment thread.');
    for (const post of thread.entries)
      if (
        !post ||
        typeof post.id !== 'string' ||
        !validId(post.author) ||
        typeof post.text !== 'string' ||
        post.text.length > 50_000 ||
        !validTime(post.createdAt) ||
        (post.editedAt !== undefined && !validTime(post.editedAt))
      )
        throw new TypeError('Invalid comment entry.');
    for (const post of thread.entries)
      if (
        post.reactions !== undefined &&
        (!post.reactions ||
          typeof post.reactions !== 'object' ||
          Array.isArray(post.reactions) ||
          Object.entries(post.reactions).some(
            ([emoji, users]) =>
              !emoji || emoji.length > 128 || !Array.isArray(users) || !unique(users),
          ))
      )
        throw new TypeError('Invalid comment reactions.');
  }
  if (workbook.suggestions !== undefined && !Array.isArray(workbook.suggestions))
    throw new TypeError('Invalid suggestions.');
  if (
    (workbook.suggestions?.length ?? 0) > 10_000 ||
    !unique((workbook.suggestions ?? []).map((item) => item?.id))
  )
    throw new TypeError('Invalid suggestion ids or count.');
  for (const suggestion of workbook.suggestions ?? []) {
    if (
      !suggestion ||
      typeof suggestion.id !== 'string' ||
      !validActor(suggestion.author) ||
      !validTime(suggestion.createdAt) ||
      (suggestion.updatedAt !== undefined && !validTime(suggestion.updatedAt)) ||
      (suggestion.commentThreadId !== undefined &&
        (typeof suggestion.commentThreadId !== 'string' || !suggestion.commentThreadId)) ||
      !['pending', 'accepted', 'rejected'].includes(suggestion.status) ||
      typeof suggestion.summary !== 'string' ||
      !Array.isArray(suggestion.changes) ||
      suggestion.changes.length > 50_000
    )
      throw new TypeError('Invalid suggestion.');
    for (const change of suggestion.changes)
      if (
        !change ||
        typeof change.sheetId !== 'string' ||
        !parseSpreadsheetAddress(change.address) ||
        !validInput(change.before) ||
        !validInput(change.after) ||
        (change.beforeStyle !== undefined && typeof change.beforeStyle !== 'string') ||
        (change.afterStyle !== undefined && typeof change.afterStyle !== 'string')
      )
        throw new TypeError('Invalid suggested cell.');
    for (const change of suggestion.changes)
      for (const field of ['beforeInputValue', 'afterInputValue'] as const)
        if (change[field] !== undefined) {
          assertNumericValue(change[field], 10_000);
          const input = field === 'beforeInputValue' ? change.before : change.after;
          if (typeof input === 'string' && input.trimStart().startsWith('='))
            throw new TypeError('A suggested canonical input cannot override a formula.');
        }
    for (const change of suggestion.changes)
      for (const field of ['beforeFormat', 'afterFormat'] as const)
        if (change[field] !== undefined) assertCellFormat(change[field]);
    if (suggestion.structure) {
      assertWorkbook({ schemaVersion: 1, worksheets: suggestion.structure.before });
      assertWorkbook({ schemaVersion: 1, worksheets: suggestion.structure.after });
      const operation = suggestion.structure.operation;
      if (
        operation &&
        (typeof operation.sheetId !== 'string' ||
          !operation.change ||
          !['row', 'column'].includes(operation.change.axis) ||
          !['insert', 'delete'].includes(operation.change.kind) ||
          !Number.isSafeInteger(operation.change.index) ||
          operation.change.index < 0 ||
          !Number.isSafeInteger(operation.change.count) ||
          operation.change.count < 1)
      )
        throw new TypeError('Invalid suggested structure operation.');
    }
  }
}

export function findOpenComment(
  workbook: Workbook,
  sheetId: string,
  range: string,
): SpreadsheetComment | undefined {
  const target = parseSpreadsheetRange(range);
  if (!target) return;
  return workbook.comments?.find((thread) => {
    if (
      thread.state === 'resolved' ||
      thread.orphaned ||
      thread.suggestion ||
      thread.sheetId !== sheetId
    )
      return false;
    const anchor = parseSpreadsheetRange(thread.range ?? '');
    return (
      !!anchor &&
      anchor.start.x <= target.end.x &&
      anchor.end.x >= target.start.x &&
      anchor.start.y <= target.end.y &&
      anchor.end.y >= target.start.y
    );
  });
}
export function addWorkbookComment(workbook: Workbook, text: string, stamp: ReviewStamp): Workbook {
  if (!text.trim()) throw new TypeError('A comment needs text.');
  const next = copyWorkbook(workbook);
  (next.comments ??= []).push({
    id: stamp.id,
    scope: 'workbook',
    state: 'open',
    entries: [{ id: stamp.id, author: stamp.actor.id, text: text.trim(), createdAt: stamp.time }],
  });
  return next;
}
/** A reviewer's local display is independent of the canonical workbook and other reviewers. */
export function previewSuggestions(workbook: Workbook, actorId: string): Workbook {
  const own = workbook.suggestions?.filter(
    (item) =>
      item.status === 'pending' && !item.outdated && !item.structure && item.author.id === actorId,
  );
  if (!own?.length) return workbook;
  const next = copyWorkbook(workbook);
  for (const item of own)
    for (const change of item.changes) {
      const sheet = next.worksheets.find((sheet) => sheet.id === change.sheetId),
        cell = parseSpreadsheetAddress(change.address);
      if (!sheet || !cell) continue;
      while (sheet.data.length <= cell.y) sheet.data.push([]);
      while (sheet.data[cell.y].length <= cell.x) sheet.data[cell.y].push(null);
      sheet.data[cell.y][cell.x] = change.after;
      if (change.afterInputValue !== undefined)
        (sheet.inputValues ??= {})[change.address] = structuredClone(change.afterInputValue);
      else if (sheet.inputValues) delete sheet.inputValues[change.address];
      if (change.afterStyle !== undefined) (sheet.style ??= {})[change.address] = change.afterStyle;
      else if (sheet.style) delete sheet.style[change.address];
      if (change.afterFormat !== undefined)
        (sheet.formats ??= {})[change.address] = structuredClone(change.afterFormat);
      else if (sheet.formats) delete sheet.formats[change.address];
    }
  return next;
}
