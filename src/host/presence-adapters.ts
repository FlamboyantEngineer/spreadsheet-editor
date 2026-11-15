// Presence without a shared room: the host carries each reader's place in a workbook to the others.
// A peer is shaped like the document editor's, with its place given as a worksheet and a range.

/** Whether a person is working: `active`, `idle` after a while without input, `away` while hidden. */
export type PresenceStatus = 'active' | 'idle' | 'away';

/** Where a person is: a worksheet, by id, and a range on it in A1 notation, such as `B2:C4`. */
export interface PresencePlace {
  readonly sheetId: string;
  readonly range: string;
}

/** What the reader shares about themselves; the host adds who they are. */
export interface LocalPresence {
  /** Their selection; null for a reader with none, such as one in a panel. */
  readonly at: PresencePlace | null;
  /** The range they are writing in now, which others may not change until it is released. */
  readonly editLock?: PresencePlace | null | undefined;
  readonly status: PresenceStatus;
}

/** Another person in the workbook. */
export interface SpreadsheetPeer extends LocalPresence {
  /** Stable, such as the person's id in the host. */
  readonly id: string;
  readonly name: string;
  /** Their picture, for the avatars. */
  readonly avatar?: string | undefined;
  /** Their colour, as a CSS colour; a colour slot from their id otherwise. */
  readonly color?: string | undefined;
}

export interface PresenceAdapter {
  /** Shares the reader's presence in a workbook, or withdraws it with null. */
  publish(workbookId: string, presence: LocalPresence | null): void;
  /**
   * Calls `onPeers` with the other people in the workbook whenever they change, without the
   * reader, until the returned function is called.
   */
  subscribe(workbookId: string, onPeers: (peers: readonly SpreadsheetPeer[]) => void): () => void;
}
