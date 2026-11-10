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
export interface CommentPost {
  id: string;
  author: ReviewActor;
  text: string;
  createdAt: string;
  updatedAt?: string;
  reactions?: Record<string, string[]>;
}
export interface SpreadsheetComment {
  id: string;
  sheetId: string;
  range: string;
  resolved: boolean;
  orphaned?: boolean;
  posts: CommentPost[];
}
export interface SuggestedCellChange {
  sheetId: string;
  address: string;
  before: CellInput;
  after: CellInput;
  beforeStyle?: string;
  afterStyle?: string;
  beforeFormat?: CellFormat;
  afterFormat?: CellFormat;
}
export interface SpreadsheetSuggestion {
  id: string;
  author: ReviewActor;
  createdAt: string;
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
  kind: 'general' | 'number' | 'percent' | 'currency' | 'scientific' | 'fraction';
  decimals?: number;
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
  const next = copyWorkbook(workbook);
  if (!text.trim()) throw new TypeError('A comment needs text.');
  next.comments = [
    ...(next.comments ?? []),
    {
      id: stamp.id,
      sheetId,
      range: parseSpreadsheetRange(range)!.address,
      resolved: false,
      posts: [
        { id: stamp.id, author: { ...stamp.actor }, text: text.trim(), createdAt: stamp.time },
      ],
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

export function replyToComment(
  workbook: Workbook,
  id: string,
  text: string,
  stamp: ReviewStamp,
): Workbook {
  if (!text.trim()) throw new TypeError('A reply needs text.');
  return updateComment(workbook, id, (thread) => {
    if (thread.resolved) throw new Error('Reopen the thread before replying.');
    thread.posts.push({
      id: stamp.id,
      author: { ...stamp.actor },
      text: text.trim(),
      createdAt: stamp.time,
    });
  });
}

export function editCommentPost(
  workbook: Workbook,
  threadId: string,
  postId: string,
  text: string,
  actorId: string,
  time: string,
): Workbook {
  if (!text.trim()) throw new TypeError('A comment needs text.');
  return updateComment(workbook, threadId, (thread) => {
    const post = thread.posts.find((post) => post.id === postId);
    if (!post || post.author.id !== actorId)
      throw new Error('Only the author can edit this comment.');
    post.text = text.trim();
    post.updatedAt = time;
  });
}

export function deleteCommentPost(
  workbook: Workbook,
  threadId: string,
  postId: string,
  actorId: string,
): Workbook {
  const next = updateComment(workbook, threadId, (thread) => {
    const post = thread.posts.find((post) => post.id === postId);
    if (!post || post.author.id !== actorId)
      throw new Error('Only the author can delete this comment.');
    thread.posts = thread.posts.filter((post) => post.id !== postId);
  });
  next.comments = next.comments?.filter((thread) => thread.posts.length);
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
    const post = thread.posts.find((post) => post.id === postId);
    if (!post) throw new RangeError('Unknown comment post.');
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
    return { ...shape, data: shape.data.map((row) => row.length) };
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
      ]);
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
          JSON.stringify(beforeFormat) !== JSON.stringify(afterFormat)
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
          });
      }
    });
  if (!structural && !changes.length) return workbook;
  next.suggestions = [
    ...(next.suggestions ?? []),
    {
      id: stamp.id,
      author: { ...stamp.actor },
      createdAt: stamp.time,
      status: 'pending',
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
          if (!ids.has(thread.sheetId)) thread.orphaned = true;
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
          JSON.stringify(sheet.formats?.[change.address]) !== JSON.stringify(change.beforeFormat)
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
        sheet.style = { ...sheet.style };
        sheet.formats = { ...sheet.formats };
        if (change.afterStyle === undefined) delete sheet.style[change.address];
        else sheet.style[change.address] = change.afterStyle;
        if (change.afterFormat === undefined) delete sheet.formats[change.address];
        else sheet.formats[change.address] = change.afterFormat;
      }
    }
  }
  suggestion.status = decision;
  suggestion.decidedBy = { ...stamp.actor };
  suggestion.decidedAt = stamp.time;
  return next;
}

export function assertReviewData(workbook: Workbook): void {
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
      typeof thread.sheetId !== 'string' ||
      !parseSpreadsheetRange(thread.range) ||
      typeof thread.resolved !== 'boolean' ||
      !Array.isArray(thread.posts) ||
      !thread.posts.length ||
      thread.posts.length > 10_000 ||
      !unique(thread.posts.map((post) => post?.id))
    )
      throw new TypeError('Invalid comment thread.');
    for (const post of thread.posts)
      if (
        !post ||
        typeof post.id !== 'string' ||
        !validActor(post.author) ||
        typeof post.text !== 'string' ||
        post.text.length > 50_000 ||
        !validTime(post.createdAt)
      )
        throw new TypeError('Invalid comment post.');
    for (const post of thread.posts)
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
