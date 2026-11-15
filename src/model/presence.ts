import type { PresencePlace, PresenceStatus, SpreadsheetPeer } from '../host/presence-adapters.js';
import { type CellSelection, normalizeSelection } from './editing.js';
import { parseSpreadsheetAddress, parseSpreadsheetRange } from './spreadsheet-addressing.js';
import type { Workbook } from './workbook.js';

/** A peer placed on the workbook: their selection and lock as cell rectangles on a worksheet. */
export interface SpreadsheetCollaborator {
  id: string;
  name: string;
  avatar?: string | undefined;
  color?: string | undefined;
  status: PresenceStatus;
  sheetId?: string;
  selection?: CellSelection;
  editLock?: CellSelection;
}
function placeSelection(workbook: Workbook, place: PresencePlace | null | undefined) {
  if (!place || !workbook.worksheets.some((sheet) => sheet.id === place.sheetId)) return undefined;
  const range = parseSpreadsheetRange(place.range);
  return range
    ? {
        sheetId: place.sheetId,
        selection: { x1: range.start.x, y1: range.start.y, x2: range.end.x, y2: range.end.y },
      }
    : undefined;
}
/**
 * The other people in the workbook, placed on its worksheets: the reader and peers without an id
 * or a name are left out, and a place on a worksheet the workbook lacks or in a malformed range is
 * dropped rather than trusted (SE-ADAPTERS-4.3).
 */
export function placePeers(
  workbook: Workbook,
  peers: readonly SpreadsheetPeer[],
  localId: string,
): SpreadsheetCollaborator[] {
  const people = new Map<string, SpreadsheetCollaborator>();
  for (const peer of peers) {
    if (peer.id === localId || !peer.id || !peer.name) continue;
    const at = placeSelection(workbook, peer.at),
      lock = placeSelection(workbook, peer.editLock);
    people.set(peer.id, {
      id: peer.id,
      name: peer.name,
      avatar: peer.avatar,
      color: peer.color,
      status: peer.status,
      sheetId: at?.sheetId ?? lock?.sheetId,
      selection: at?.selection,
      editLock: lock && (!at || lock.sheetId === at.sheetId) ? lock.selection : undefined,
    });
  }
  return [...people.values()].sort((a, b) => a.name.localeCompare(b.name));
}
export function collaboratorSheet(workbook: Workbook, person: SpreadsheetCollaborator) {
  return workbook.worksheets.find((sheet) => sheet.id === person.sheetId);
}

/** A peer's edit lock as a normalized selection; a malformed lock locks nothing. */
function lockRange(peer: SpreadsheetCollaborator): CellSelection | undefined {
  if (!peer.editLock) return undefined;
  try {
    return normalizeSelection(peer.editLock);
  } catch {
    return undefined;
  }
}

/** The peer whose edit lock overlaps a selection on a worksheet. */
export function selectionLock(
  workbook: Workbook,
  { sheetId, selection }: { sheetId: string; selection: CellSelection },
  peers: readonly SpreadsheetCollaborator[],
): SpreadsheetCollaborator | undefined {
  return peers.find((peer) => {
    if (!peer.editLock || collaboratorSheet(workbook, peer)?.id !== sheetId) return false;
    const lock = lockRange(peer);
    return Boolean(
      lock &&
      lock.x1 <= selection.x2 &&
      lock.x2 >= selection.x1 &&
      lock.y1 <= selection.y2 &&
      lock.y2 >= selection.y1,
    );
  });
}

type Sheet = Workbook['worksheets'][number];

function dataChanged(sheet: Sheet, next: Sheet, range: CellSelection): boolean {
  const lastRow = Math.min(range.y2, Math.max(sheet.data.length, next.data.length) - 1);
  for (let y = range.y1; y <= lastRow; y++) {
    const width = Math.max(sheet.data[y]?.length ?? 0, next.data[y]?.length ?? 0);
    for (let x = range.x1; x <= Math.min(range.x2, width - 1); x++)
      if ((sheet.data[y]?.[x] ?? null) !== (next.data[y]?.[x] ?? null)) return true;
  }
  return false;
}

function decorationChanged(sheet: Sheet, next: Sheet, range: CellSelection): boolean {
  const changed = new Set([
    ...Object.keys(sheet.style ?? {}),
    ...Object.keys(next.style ?? {}),
    ...Object.keys(sheet.formats ?? {}),
    ...Object.keys(next.formats ?? {}),
  ]);
  return [...changed].some((address) => {
    const cell = parseSpreadsheetAddress(address);
    return Boolean(
      cell &&
      cell.x >= range.x1 &&
      cell.x <= range.x2 &&
      cell.y >= range.y1 &&
      cell.y <= range.y2 &&
      (sheet.style?.[address] !== next.style?.[address] ||
        JSON.stringify(sheet.formats?.[address]) !== JSON.stringify(next.formats?.[address])),
    );
  });
}

/** Advisory UI conflict detection. Authorization and durable conflict handling remain host-owned. */
export function lockedChange(
  before: Workbook,
  after: Workbook,
  peers: readonly SpreadsheetCollaborator[],
): SpreadsheetCollaborator | undefined {
  return peers.find((peer) => {
    const range = lockRange(peer);
    if (!range) return false;
    const sheet = collaboratorSheet(before, peer),
      next = after.worksheets.find((value) => value.id === sheet?.id);
    if (!sheet || !next) return false;
    return dataChanged(sheet, next, range) || decorationChanged(sheet, next, range);
  });
}
