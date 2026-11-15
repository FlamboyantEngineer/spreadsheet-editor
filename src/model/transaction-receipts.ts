/** Session-local causal coverage; owns no saving, scheduling, provider or UI state. */
export function createTransactionReceiptTracker() {
  const pending = new Map<string, number>();
  let latestRevision = 0;
  function forgetThrough(revision: number): void {
    for (const [id, accepted] of pending) if (accepted <= revision) pending.delete(id);
  }
  return {
    record(transactionId: string, localRevision: number): void {
      if (
        !transactionId ||
        pending.has(transactionId) ||
        !Number.isSafeInteger(localRevision) ||
        localRevision <= latestRevision
      )
        throw new TypeError(
          'Accepted transactions require unique IDs and increasing positive revisions.',
        );
      pending.set(transactionId, localRevision);
      latestRevision = localRevision;
    },
    acknowledge(transactionId: string): number | undefined {
      const revision = pending.get(transactionId);
      if (revision !== undefined) forgetThrough(revision);
      return revision;
    },
    /** A successful existing Save/Flush receipt also releases covered transaction IDs. */
    forgetThrough,
  };
}
