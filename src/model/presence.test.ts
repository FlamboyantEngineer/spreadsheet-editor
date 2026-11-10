import { describe, expect, it } from 'vitest';
import { createWorkbook } from './workbook.js';
import { collaboratorSheet, remoteCollaborators, selectionLock } from './presence.js';
describe('Host SE-CO: host-neutral presence', () => {
  it('deduplicates remote identities, removes self and resolves sheets without trusting peer coordinates', () => {
    const workbook = createWorkbook();
    workbook.worksheets[0].name = 'Mass Budget';
    const peers = remoteCollaborators(
      [
        { id: 'self', name: 'Me' },
        { id: 'peer', name: 'Old', updatedAt: 1 },
        {
          id: 'peer',
          name: 'Morgan',
          updatedAt: 2,
          sheetName: '  mass   budget ',
          editLock: { x1: 0, x2: 1, y1: 0, y2: 0 },
        },
      ],
      'self',
    );
    expect(peers).toHaveLength(1);
    expect(collaboratorSheet(workbook, peers[0])?.id).toBe('sheet-1');
    expect(selectionLock(workbook, 'sheet-1', { x1: 1, x2: 1, y1: 0, y2: 0 }, peers)?.name).toBe(
      'Morgan',
    );
    expect(
      selectionLock(workbook, 'sheet-1', { x1: 2, x2: 2, y1: 0, y2: 0 }, peers),
    ).toBeUndefined();
    peers[0].editLock!.x1 = -1;
    expect(
      selectionLock(workbook, 'sheet-1', { x1: 0, x2: 0, y1: 0, y2: 0 }, peers),
    ).toBeUndefined();
  });
});
