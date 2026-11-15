import { getColorIndex } from '@flamboyantengineer/ui-components/core';

/** How many colour slots a collaborator is given one of. */
const COLLABORATOR_COLOR_SLOTS = 8;

/**
 * A collaborator's colour slot from their id, the base library's `getColorIndex`, so a person has
 * the slot the base library's avatars and the other editors give the same id (SE-ADAPTERS-4.3).
 */
export function collaboratorColorIndex(id: string): number {
  return getColorIndex(id, COLLABORATOR_COLOR_SLOTS);
}
