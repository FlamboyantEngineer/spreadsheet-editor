'use client';

import { Banner } from '@flamboyantengineer/ui-components/banner';
import { Button } from '@flamboyantengineer/ui-components/button';
import { InlineError } from '@flamboyantengineer/ui-components/inline-error';
import { VisuallyHidden } from '@flamboyantengineer/ui-components/visually-hidden';
import type { WorkbookVersion } from '../host/adapters.js';
import type { SpreadsheetEditorController } from './use-editor.js';

type EditorProps = { editor: SpreadsheetEditorController };

/** Says which saved version the reader is viewing, with restore and return actions. */
export function HistoricalBanner({
  editor: e,
  version,
}: EditorProps & { version: WorkbookVersion }) {
  return (
    <Banner
      className="se-editor__historical"
      tone="info"
      description={
        <span>
          You are viewing {version.label ? <q>{version.label}</q> : 'a version'}, saved{' '}
          {new Intl.DateTimeFormat(undefined, {
            dateStyle: 'medium',
            timeStyle: 'long',
          }).format(new Date(version.savedAt))}
          {version.author && <> by {e.personName(version.author)}</>}. It is read-only.
        </span>
      }
      actions={
        <>
          {e.canRestoreVersion && (
            <Button size="sm" onPress={() => void e.restoreVersion(version)}>
              Restore this version
            </Button>
          )}
          <Button size="sm" onPress={e.returnToCurrent}>
            View current version
          </Button>
        </>
      }
    />
  );
}

/** The calculation status and the selection summary, for assistive technology. */
export function CalculationStatus({ editor: e }: EditorProps) {
  const mode = e.props.calculation?.mode;
  return (
    <VisuallyHidden>
      <p className="se-editor__status" aria-label="Calculation status" role="status">
        {e.calculation.error ??
          (e.calculation.pending
            ? 'Calculating…'
            : mode === 'external'
              ? 'Published by host'
              : `Calculated ${mode === 'server' ? 'on server' : 'on client'}`)}
      </p>
      <span aria-label="Selection summary">
        Count: {e.summary.count} · Sum: {e.summary.sum} · Average: {e.summary.average}
      </span>
    </VisuallyHidden>
  );
}

/** A save conflict's choice between the reader's changes and the saved workbook. */
function ConflictBanner({ editor: e }: EditorProps) {
  return (
    <Banner
      className="se-editor__conflict"
      tone="warning"
      description="Someone else saved this workbook after you opened it. Keep your changes over theirs, or take the saved workbook and drop yours."
      actions={
        <>
          <Button size="sm" onPress={() => void e.resolveConflict('mine')}>
            Keep my changes
          </Button>
          <Button size="sm" onPress={() => void e.resolveConflict('stored')}>
            Use saved workbook
          </Button>
        </>
      }
    />
  );
}

/** Offers the unsaved changes the session kept from an earlier visit. */
function DraftBanner({ editor: e, stale }: EditorProps & { stale: boolean }) {
  return (
    <Banner
      className="se-editor__draft"
      tone="info"
      description={
        stale
          ? 'Unsaved changes from an earlier visit were kept, but the workbook was saved since. Restoring them replaces the saved workbook when you save.'
          : 'Unsaved changes from an earlier visit were kept.'
      }
      actions={
        <>
          <Button size="sm" onPress={() => e.props.session?.restoreDraft()}>
            Restore changes
          </Button>
          <Button size="sm" onPress={() => e.props.session?.discardDraft()}>
            Discard
          </Button>
        </>
      }
    />
  );
}

/** The calculation error, the session's conflict and kept draft, and the editor's notice. */
export function EditorNotices({ editor: e }: EditorProps) {
  return (
    <>
      {e.calculation.error && (
        <InlineError className="se-editor__notice">{e.calculation.error}</InlineError>
      )}
      {e.sessionStatus?.phase === 'conflict' && <ConflictBanner editor={e} />}
      {e.sessionDraft && <DraftBanner editor={e} stale={e.sessionDraft.stale} />}
      {e.notice && (
        <Banner
          className="se-editor__notice"
          tone={e.noticeTone}
          key={e.noticeRevision}
          description={e.notice}
          onDismiss={() => e.setNotice(null)}
          dismissLabel="Dismiss"
        />
      )}
    </>
  );
}

/** The host's insertion commands for the selected cells. */
export function InsertionActions({ editor: e }: EditorProps) {
  return (
    <div
      className="se-editor__insertion"
      role="group"
      aria-label="Insert selected cells"
      onFocusCapture={() => e.grid.current?.releaseKeyboard()}
    >
      <Button
        onPress={() => e.attempt(() => e.insert('table'))}
        isDisabled={e.draftDirty || e.calculation.pending || !!e.calculation.error}
      >
        Insert as document table
      </Button>
      <Button onPress={() => e.attempt(() => e.insert('spreadsheet'))}>
        Insert as live spreadsheet
      </Button>
    </div>
  );
}
