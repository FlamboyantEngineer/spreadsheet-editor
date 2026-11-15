import { getColorIndex } from '@flamboyantengineer/ui-components/core';
import { describe, expect, it } from 'vitest';
import { collaboratorColorIndex } from './collaborator-color.js';

describe('SE-ADAPTERS-4.3: collaborator colour slots', () => {
  it('SE-ADAPTERS-4.3: gives a collaborator the slot the base library gives their id', () => {
    for (const id of ['ada', 'grace', 'hedy', 'p1', ' p9 ', 'name-with-é']) {
      expect(collaboratorColorIndex(id)).toBe(getColorIndex(id, 8));
    }
  });
});
