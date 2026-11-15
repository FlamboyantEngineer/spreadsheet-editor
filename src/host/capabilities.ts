/**
 * What the current reader may do with one workbook, as the host's permission system grants it.
 * The editor never reads roles: the host maps its own roles to these capabilities.
 */
export interface EditorCapabilities {
  /** Change cells and structure directly. */
  edit: boolean;
  /** Propose changes as suggestions that another reader accepts or rejects. */
  suggest: boolean;
  /** Add, reply to and react to comments. */
  comment: boolean;
  /** Accept and reject other readers' suggestions. */
  decide: boolean;
  /** Save a named version of the workbook. */
  saveVersion: boolean;
  /** Open the workbook's saved versions. */
  history: boolean;
}

/** How the editor treats input: authoring, reviewing (suggesting or commenting) or reading. */
export type SpreadsheetMode = 'author' | 'review' | 'read';

export interface SpreadsheetModeContext {
  /** A historical version is open; it never accepts changes. */
  historical?: boolean | undefined;
}

/** No capabilities: the workbook is read-only. */
export const READ_ONLY_CAPABILITIES: EditorCapabilities = Object.freeze({
  edit: false,
  suggest: false,
  comment: false,
  decide: false,
  saveVersion: false,
  history: false,
});

/** Every capability, for a reader who owns the workbook. */
export const FULL_CAPABILITIES: EditorCapabilities = Object.freeze({
  edit: true,
  suggest: true,
  comment: true,
  decide: true,
  saveVersion: true,
  history: true,
});

/**
 * The mode a reader's capabilities allow at most: `read` for a historical version, `author` for a
 * reader who may edit, `review` for one who may only suggest or comment, and `read` otherwise.
 */
export function resolveSpreadsheetMode(
  capabilities: EditorCapabilities,
  context: SpreadsheetModeContext = {},
): SpreadsheetMode {
  if (context.historical) return 'read';
  if (capabilities.edit) return 'author';
  if (capabilities.suggest || capabilities.comment) return 'review';
  return 'read';
}
