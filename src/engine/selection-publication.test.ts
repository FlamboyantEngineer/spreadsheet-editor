import { describe, expect, it, vi } from 'vitest';
import { createSelectionPublication } from './selection-publication.js';

describe('SE-UI-1: bounded selection publication', () => {
  it('coalesces pointer bursts once per frame and publishes final selection', () => {
    const frames = new Map<number, FrameRequestCallback>();
    let id = 0;
    const publish = vi.fn();
    const selections = createSelectionPublication(
      publish,
      (callback) => {
        frames.set(++id, callback);
        return id;
      },
      (request) => frames.delete(request),
    );
    const range = (x: number) => ({ x1: 0, y1: 0, x2: x, y2: 99 });
    selections.start();
    for (let x = 0; x < 26; x++) selections.update(range(x));
    expect(publish).not.toHaveBeenCalled();
    expect(frames.size).toBe(1);
    frames.values().next().value!(0);
    expect(publish).toHaveBeenCalledExactlyOnceWith(range(25));
    selections.update(range(24));
    selections.finish();
    expect(publish).toHaveBeenCalledTimes(2);
    expect(publish).toHaveBeenLastCalledWith(range(24));
    selections.update(range(1));
    expect(publish).toHaveBeenCalledTimes(3);
    selections.dispose();
  });

  it('publishes the final selection after its last animation frame and cancels on disposal', () => {
    let callback: FrameRequestCallback = () => {};
    const publish = vi.fn(),
      cancel = vi.fn();
    const selections = createSelectionPublication(
      publish,
      (next) => ((callback = next), 1),
      cancel,
    );
    const range = { x1: 0, y1: 0, x2: 25, y2: 99 };
    selections.start();
    selections.update(range);
    callback(0);
    selections.finish();
    expect(publish).toHaveBeenCalledTimes(2);
    expect(publish).toHaveBeenLastCalledWith(range);
    selections.start();
    selections.update(range);
    selections.dispose();
    expect(cancel).toHaveBeenCalledWith(1);
    expect(publish).toHaveBeenCalledTimes(2);
  });
});
