import { describe, expect, it } from 'vitest';
import { createTransactionReceiptTracker } from './transaction-receipts.js';

describe('SE-HOST-5.1: causal receipt bookkeeping for host sessions', () => {
  it('cannot infer later-edit coverage from an earlier or unknown receipt', () => {
    const tracker = createTransactionReceiptTracker();
    tracker.record('first', 1);
    tracker.record('second', 2);
    expect(tracker.acknowledge('another-session')).toBeUndefined();
    expect(tracker.acknowledge('first')).toBe(1);
    expect(tracker.acknowledge('first')).toBeUndefined();
    expect(tracker.acknowledge('second')).toBe(2);
  });
  it('releases only submitted Save/Flush coverage and includes comment-only transactions', () => {
    const tracker = createTransactionReceiptTracker();
    tracker.record('cell-edit', 1);
    const submittedRevision = 1;
    tracker.record('comment-reply', 2);
    tracker.record('review-decision', 3);
    tracker.forgetThrough(submittedRevision);
    expect(tracker.acknowledge('cell-edit')).toBeUndefined();
    expect(tracker.acknowledge('comment-reply')).toBe(2);
    expect(tracker.acknowledge('review-decision')).toBe(3);
  });
  it('a later cumulative acknowledgement releases every earlier pending receipt', () => {
    const tracker = createTransactionReceiptTracker();
    tracker.record('one', 1);
    tracker.record('two', 2);
    tracker.record('three', 3);
    expect(tracker.acknowledge('three')).toBe(3);
    expect(tracker.acknowledge('one')).toBeUndefined();
    expect(tracker.acknowledge('two')).toBeUndefined();
  });
  it('rejects duplicate pending IDs and non-causal revision registration without changing coverage', () => {
    const tracker = createTransactionReceiptTracker();
    tracker.record('accepted', 1);
    expect(() => tracker.record('accepted', 2)).toThrow('unique IDs');
    expect(() => tracker.record('later', 1)).toThrow('increasing');
    expect(tracker.acknowledge('accepted')).toBe(1);
    tracker.record('later', 2);
    expect(tracker.acknowledge('later')).toBe(2);
  });
});
