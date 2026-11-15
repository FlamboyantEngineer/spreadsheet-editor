import type { ReviewActor } from './review.js';
import { worksheetNameKey, type Workbook } from './workbook.js';
import { normalizeSelection, type CellSelection } from './editing.js';
import { parseSpreadsheetAddress } from './spreadsheet-addressing.js';

export interface SpreadsheetCollaborator extends ReviewActor {
  sheetId?: string;
  sheetName?: string;
  sheetIndex?: number;
  selection?: CellSelection;
  editLock?: CellSelection;
  activity?: 'Editing' | 'Reviewing' | 'Viewing' | 'Idle';
  updatedAt?: number;
}
export interface SpreadsheetPresence {
  sheetId: string;
  selection: CellSelection;
  editLock?: CellSelection;
  activity: 'Editing' | 'Reviewing' | 'Viewing';
  timestamp: number;
}
export function collaboratorColorIndex(id: string): number {
  let hash = 0;
  for (const character of id) hash = (hash * 31 + character.charCodeAt(0)) >>> 0;
  return hash % 8;
}
export function remoteCollaborators(
  peers: readonly SpreadsheetCollaborator[],
  localId: string,
): SpreadsheetCollaborator[] {
  const people = new Map<string, SpreadsheetCollaborator>();
  for (const person of peers) {
    if (person.id === localId || !person.id || !person.name) continue;
    const previous = people.get(person.id);
    if (!previous || (person.updatedAt ?? 0) >= (previous.updatedAt ?? 0))
      people.set(person.id, person);
  }
  return [...people.values()].sort((a, b) => a.name.localeCompare(b.name));
}
export function collaboratorSheet(workbook: Workbook, person: SpreadsheetCollaborator) {
  return (
    workbook.worksheets.find((sheet) => sheet.id === person.sheetId) ??
    workbook.worksheets.find(
      (sheet) =>
        person.sheetName !== undefined &&
        worksheetNameKey(sheet.name) === worksheetNameKey(person.sheetName),
    ) ??
    (person.sheetIndex === undefined ? undefined : workbook.worksheets[person.sheetIndex])
  );
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
