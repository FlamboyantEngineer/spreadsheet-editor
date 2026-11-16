import type { NumericValue } from '../calculation/values.js';
import type { StructureChange } from './structure.js';
import type { CellInput, Worksheet } from './workbook.js';

export {
  addComment,
  addWorkbookComment,
  type CommentEntryRef,
  deleteCommentEntry,
  editCommentEntry,
  findOpenComment,
  type NewComment,
  replyToComment,
  setCommentState,
  toggleCommentReaction,
  updateComment,
} from './review-comments.js';
export {
  decideSuggestion,
  previewSuggestions,
  proposeWorkbookChange,
  type SuggestionDecision,
  type WorkbookProposal,
} from './review-suggestions.js';

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
