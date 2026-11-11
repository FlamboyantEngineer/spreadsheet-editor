import type { JSpreadsheetSelectionRange } from './jspreadsheet-host.js';

/** Pointer drags publish a bounded amount of React work; keyboard navigation is immediate. */
export function createSelectionPublication(
  publish: (range: JSpreadsheetSelectionRange) => void,
  request: (callback: FrameRequestCallback) => number = requestAnimationFrame,
  cancel: (request: number) => void = cancelAnimationFrame,
) {
  let dragging = false;
  let pending: JSpreadsheetSelectionRange | null = null;
  let last: JSpreadsheetSelectionRange | null = null;
  let frame = 0;
  function flush() {
    if (frame) cancel(frame);
    frame = 0;
    const range = pending;
    pending = null;
    if (range) publish(range);
  }
  return {
    start() {
      dragging = true;
      last = null;
    },
    update(range: JSpreadsheetSelectionRange) {
      if (!dragging) {
        publish(range);
        return;
      }
      last = range;
      pending = range;
      if (!frame) frame = request(flush);
    },
    finish() {
      // Also used for cancellation, lost capture and window blur. A final delivered
      // vendor selection must never disappear just because its frame was cancelled.
      if (dragging) pending = last;
      dragging = false;
      flush();
      last = null;
    },
    dispose() {
      if (frame) cancel(frame);
      frame = 0;
      pending = null;
      last = null;
      dragging = false;
    },
  };
}
