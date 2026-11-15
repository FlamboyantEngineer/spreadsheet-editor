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
export function collaboratorColorIndex(id: string): number {
  let hash = 0;
  for (const character of id) hash = (hash * 31 + character.charCodeAt(0)) >>> 0;
  return hash % 8;
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

export function selectionLock(
  workbook: Workbook,
  sheetId: string,
  selection: CellSelection,
  peers: readonly SpreadsheetCollaborator[],
): SpreadsheetCollaborator | undefined {
  return peers.find((peer) => {
    if (!peer.editLock || collaboratorSheet(workbook, peer)?.id !== sheetId) return false;
    try {
      const lock = normalizeSelection(peer.editLock);
      return (
        lock.x1 <= selection.x2 &&
        lock.x2 >= selection.x1 &&
        lock.y1 <= selection.y2 &&
        lock.y2 >= selection.y1
      );
    } catch {
      return false;
    }
  });
}
/** Advisory UI conflict detection. Authorization and durable conflict handling remain host-owned. */
export function lockedChange(
  before: Workbook,
  after: Workbook,
  peers: readonly SpreadsheetCollaborator[],
): SpreadsheetCollaborator | undefined {
  for (const peer of peers) {
    if (!peer.editLock) continue;
    let range: CellSelection;
    try {
      range = normalizeSelection(peer.editLock);
    } catch {
      continue;
    }
    const sheet = collaboratorSheet(before, peer),
      next = after.worksheets.find((value) => value.id === sheet?.id);
    if (!sheet || !next) continue;
    const changed = new Set([
      ...Object.keys(sheet.style ?? {}),
      ...Object.keys(next.style ?? {}),
      ...Object.keys(sheet.formats ?? {}),
      ...Object.keys(next.formats ?? {}),
    ]);
    for (
      let y = range.y1;
      y <= Math.min(range.y2, Math.max(sheet.data.length, next.data.length) - 1);
      y++
    ) {
      for (
        let x = range.x1;
        x <=
        Math.min(range.x2, Math.max(sheet.data[y]?.length ?? 0, next.data[y]?.length ?? 0) - 1);
        x++
      )
        if ((sheet.data[y]?.[x] ?? null) !== (next.data[y]?.[x] ?? null)) return peer;
    }
    for (const address of changed) {
      const cell = parseSpreadsheetAddress(address);
      if (
        cell &&
        cell.x >= range.x1 &&
        cell.x <= range.x2 &&
        cell.y >= range.y1 &&
        cell.y <= range.y2 &&
        (sheet.style?.[address] !== next.style?.[address] ||
          JSON.stringify(sheet.formats?.[address]) !== JSON.stringify(next.formats?.[address]))
      )
        return peer;
    }
  }
  return undefined;
}
