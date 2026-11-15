import { describe, expect, it } from 'vitest';
import { collaboratorSheet, placePeers, selectionLock } from './presence.js';
import { createWorkbook } from './workbook.js';

describe('SE-ADAPTERS-4.3: peers placed on the workbook', () => {
  it('leaves out the reader and drops places the workbook does not have', () => {
    const workbook = createWorkbook();
    const peers = placePeers(
      workbook,
      [
        { id: 'self', name: 'Me', at: null, status: 'active' },
        {
          id: 'peer',
          name: 'Morgan',
          at: { sheetId: 'sheet-1', range: 'b1' },
          editLock: { sheetId: 'sheet-1', range: 'A1:B1' },
          status: 'idle',
        },
        { id: 'lost', name: 'Avery', at: { sheetId: 'gone', range: 'A1' }, status: 'active' },
        { id: 'bad', name: 'Blake', at: { sheetId: 'sheet-1', range: 'A0' }, status: 'away' },
        { id: '', name: 'Nobody', at: null, status: 'active' },
      ],
      'self',
    );
    expect(peers.map((peer) => peer.id)).toEqual(['lost', 'bad', 'peer']);
    const [lost, bad, peer] = peers;
    expect(lost.sheetId).toBeUndefined();
    expect(bad.selection).toBeUndefined();
    expect(peer).toMatchObject({
      status: 'idle',
      sheetId: 'sheet-1',
      selection: { x1: 1, x2: 1, y1: 0, y2: 0 },
      editLock: { x1: 0, x2: 1, y1: 0, y2: 0 },
    });
    expect(collaboratorSheet(workbook, peer)?.id).toBe('sheet-1');
    expect(selectionLock(workbook, 'sheet-1', { x1: 1, x2: 1, y1: 0, y2: 0 }, peers)?.name).toBe(
      'Morgan',
    );
    expect(
      selectionLock(workbook, 'sheet-1', { x1: 2, x2: 2, y1: 0, y2: 0 }, peers),
    ).toBeUndefined();
  });
});
