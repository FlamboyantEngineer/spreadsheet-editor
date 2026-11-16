'use client';

import {
  ClockGlyph,
  EditGlyph,
  InfoGlyph,
  ReplyGlyph,
} from '@flamboyantengineer/ui-components/core';
import { DockPanel, DockPanelPane } from '@flamboyantengineer/ui-components/dock-panel';
import { CommentsPanel } from './comments-panel.js';
import type { SpreadsheetEditorProps } from './contracts.js';
import { HistoryPanel, PropertiesPanel } from './review-panels.js';
import { SuggestionsPanel } from './suggestions-panel.js';
import type { SpreadsheetEditorController } from './use-editor.js';

const BUILT_IN_PANES = ['properties', 'history', 'comments', 'suggestions'];

/** Open threads that are not suggestion discussions, and pending suggestions. */
function reviewCounts({ workbook }: SpreadsheetEditorController) {
  const pending = workbook.suggestions?.filter((item) => item.status === 'pending').length;
  const open =
    workbook.comments?.filter((thread) => thread.state === 'open' && !thread.suggestion).length ??
    0;
  return { comments: open + (pending ?? 0), suggestions: pending };
}

/** The dock of built-in and host panels beside the grid. */
export function EditorSidebar({
  editor: e,
  panelRef,
  panelId,
  panelData,
}: { editor: SpreadsheetEditorController } & Pick<
  SpreadsheetEditorProps,
  'panelRef' | 'panelId' | 'panelData'
>) {
  const { props } = e;
  const hostPane = (id: string) => props.panels?.find((panel) => panel.id === id)?.content;
  const counts = reviewCounts(e);
  return (
    <div
      className="se-editor__sidebar"
      {...panelData}
      id={panelId}
      ref={panelRef}
      onFocusCapture={() => e.grid.current?.releaseKeyboard()}
    >
      <DockPanel
        label="Spreadsheet panels"
        selectedKey={e.panel ?? undefined}
        onSelectionChange={(key) => e.openPanel(String(key))}
        isExpanded={e.panelOpen}
        onExpandedChange={e.setPanelOpen}
        defaultSize={340}
        minSize={260}
        maxSize={600}
      >
        <DockPanelPane id="properties" label="Properties" icon={InfoGlyph}>
          {hostPane('properties') ?? <PropertiesPanel editor={e} />}
        </DockPanelPane>
        <DockPanelPane id="history" label="History" icon={ClockGlyph}>
          {hostPane('history') ?? <HistoryPanel editor={e} />}
        </DockPanelPane>
        <DockPanelPane id="comments" label="Comments" icon={ReplyGlyph} count={counts.comments}>
          {hostPane('comments') ?? <CommentsPanel editor={e} />}
        </DockPanelPane>
        <DockPanelPane
          id="suggestions"
          label="Suggestions"
          icon={EditGlyph}
          count={counts.suggestions}
        >
          <SuggestionsPanel editor={e} />
        </DockPanelPane>
        {props.panels
          ?.filter((panel) => !BUILT_IN_PANES.includes(panel.id))
          .map((panel) => (
            <DockPanelPane key={panel.id} id={panel.id} label={panel.label} icon={InfoGlyph}>
              {panel.content}
            </DockPanelPane>
          ))}
      </DockPanel>
    </div>
  );
}
