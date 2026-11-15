import type { EditorPerson } from './adapters.js';
import type { LocalPresence, PresenceAdapter, SpreadsheetPeer } from './presence-adapters.js';

/** People who share presence in memory, such as two editors on one lab page. */
export interface MemoryPresenceRoom {
  /** The presence adapter of one person, who publishes as themselves and sees everyone else. */
  join(person: EditorPerson): PresenceAdapter;
}

type Listener = { person: string; onPeers: (peers: readonly SpreadsheetPeer[]) => void };

/** Presence kept in memory and shared by everyone who joins (SE-ADAPTERS-3.4). */
export function createMemoryPresenceRoom(): MemoryPresenceRoom {
  const published = new Map<string, Map<string, SpreadsheetPeer>>();
  const listeners = new Map<string, Set<Listener>>();
  const peersFor = (workbookId: string, person: string) =>
    [...(published.get(workbookId)?.values() ?? [])]
      .filter((peer) => peer.id !== person)
      .sort((a, b) => a.name.localeCompare(b.name) || a.id.localeCompare(b.id));
  const notify = (workbookId: string) =>
    listeners.get(workbookId)?.forEach((listener) => {
      listener.onPeers(peersFor(workbookId, listener.person));
    });
  return {
    join(person) {
      return {
        publish(workbookId, presence: LocalPresence | null) {
          const people = published.get(workbookId) ?? new Map<string, SpreadsheetPeer>();
          published.set(workbookId, people);
          if (presence)
            people.set(person.id, {
              id: person.id,
              name: person.name,
              ...(person.avatarUrl === undefined ? {} : { avatar: person.avatarUrl }),
              ...structuredClone(presence),
            });
          else people.delete(person.id);
          notify(workbookId);
        },
        subscribe(workbookId, onPeers) {
          const listener: Listener = { person: person.id, onPeers };
          const set = listeners.get(workbookId) ?? new Set<Listener>();
          listeners.set(workbookId, set);
          set.add(listener);
          onPeers(peersFor(workbookId, person.id));
          return () => set.delete(listener);
        },
      };
    },
  };
}
