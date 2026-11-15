import { useCallback, useEffect, useMemo, useState } from 'react';
import type { IdentityAdapter } from '../host/adapters.js';
import type { ReviewActor, SpreadsheetComment } from '../model/review.js';

/**
 * Names the people who wrote, resolved or reacted to the workbook's comments, and the `others`
 * the editor shows, such as the authors of saved versions. Comments and versions store person ids
 * only; the host's identity adapter resolves them, and an id it does not know stands for itself
 * (SE-REVIEW-1.6).
 */
export function usePersonNames(
  identity: IdentityAdapter | undefined,
  comments: readonly SpreadsheetComment[] | undefined,
  actor: ReviewActor,
  others: readonly (string | null)[] = [],
): (id: string) => string {
  const [names, setNames] = useState<ReadonlyMap<string, string>>(new Map());
  const wanted = useMemo(() => {
    const ids = new Set<string>();
    for (const thread of comments ?? []) {
      for (const entry of thread.entries) {
        ids.add(entry.author);
        for (const people of Object.values(entry.reactions ?? {}))
          people.forEach((id) => ids.add(id));
      }
      if (thread.resolvedBy) ids.add(thread.resolvedBy);
    }
    for (const id of others) if (id) ids.add(id);
    ids.delete(actor.id);
    return [...ids].filter((id) => !names.has(id)).sort();
  }, [comments, others, actor.id, names]);
  const key = wanted.join('\n');

  useEffect(() => {
    if (!identity || !key) return;
    let current = true;
    const ids = key.split('\n');
    identity.resolvePeople(ids).then(
      (people) => {
        if (!current) return;
        setNames((known) => {
          const next = new Map(known);
          for (const id of ids) next.set(id, people.find((person) => person.id === id)?.name ?? id);
          return next;
        });
      },
      () => {
        if (current)
          setNames((known) => new Map([...known, ...ids.map((id) => [id, id] as const)]));
      },
    );
    return () => {
      current = false;
    };
  }, [identity, key]);

  return useCallback(
    (id: string) => (id === actor.id ? actor.name : (names.get(id) ?? id)),
    [actor, names],
  );
}
