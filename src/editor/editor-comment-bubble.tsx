'use client';

import { Popover } from '@flamboyantengineer/ui-components/popover';
import { useEffectEvent, useLayoutEffect, useRef, useState } from 'react';
import { findOpenComment } from '../model/review.js';
import { formatSpreadsheetAddress } from '../model/spreadsheet-addressing.js';
import { CommentsPanel } from './comments-panel.js';
import type { SpreadsheetEditorController } from './use-editor.js';

type Anchor = { left: number; top: number } | null;

/** Where the open comment's bubble points: the selection's top right corner in the grid. */
export function useCommentBubbleAnchor(e: SpreadsheetEditorController) {
  const [anchor, setAnchor] = useState<Anchor>(null);
  function positionComment() {
    if (!e.activeCommentId) return;
    const rect = e.grid.current?.getAnchor(e.selection);
    if (rect)
      setAnchor((previous) =>
        previous?.left === rect.right && previous.top === rect.top
          ? previous
          : { left: rect.right, top: rect.top },
      );
  }
  const positionCommentOnChange = useEffectEvent(positionComment);
  useLayoutEffect(() => {
    // eslint-disable-next-line react-hooks/set-state-in-effect -- the bubble is placed from the grid's committed cell geometry
    positionCommentOnChange();
  }, [e.activeCommentId, e.selection, e.workbook]);
  return { anchor, positionComment };
}

/** Opens the open thread of a cell, or the discussion of a pending suggestion on it. */
export function openCellComment(
  e: SpreadsheetEditorController,
  { x, y }: { x: number; y: number },
) {
  const address = formatSpreadsheetAddress(x, y);
  const thread = findOpenComment(e.workbook, e.sheet.id, address);
  if (thread) {
    e.openComment(thread.id);
    return;
  }
  const suggestion = e.workbook.suggestions?.find(
    (item) =>
      item.status === 'pending' &&
      !item.outdated &&
      item.changes.some((change) => change.sheetId === e.sheet.id && change.address === address),
  );
  if (suggestion) e.openSuggestion(suggestion.id);
}

/** The open comment thread or new comment, in a bubble beside its cell. */
export function CommentBubble({
  editor: e,
  anchor,
}: {
  editor: SpreadsheetEditorController;
  anchor: Anchor;
}) {
  const trigger = useRef<HTMLSpanElement | null>(null);
  return (
    <>
      <span
        ref={trigger}
        aria-hidden="true"
        style={{
          position: 'fixed',
          left: anchor?.left ?? 0,
          top: anchor?.top ?? 0,
          width: 0,
          height: 0,
          pointerEvents: 'none',
        }}
      />
      <Popover
        className="se-comment-bubble"
        triggerRef={trigger}
        placement="right top"
        offset={8}
        isNonModal
        isOpen={!!e.activeCommentId && !!anchor}
        onOpenChange={(open) => {
          if (!open) {
            e.setActiveCommentId(null);
            e.setCommentComposer(false);
          }
        }}
        title={e.activeCommentId === 'draft' ? 'New comment' : 'Comment'}
        showCloseButton
      >
        <div onFocusCapture={() => e.grid.current?.releaseKeyboard()}>
          <CommentsPanel editor={e} threadId={e.activeCommentId ?? undefined} />
        </div>
      </Popover>
    </>
  );
}
