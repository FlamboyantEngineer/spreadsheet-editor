'use client';

import { Attribution } from '@flamboyantengineer/ui-components/attribution';
import { Button } from '@flamboyantengineer/ui-components/button';
import { ContextMenuTrigger } from '@flamboyantengineer/ui-components/context-menu';
import { InlineError } from '@flamboyantengineer/ui-components/inline-error';
import { Menu } from '@flamboyantengineer/ui-components/menu';
import { Spinner } from '@flamboyantengineer/ui-components/progress-bar';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { useState } from 'react';
import type { WorkbookVersion } from '../host/adapters.js';
import { selectionAddress } from '../model/editing.js';
import { editorCommandMetadata } from './editor-commands.js';
import type { SpreadsheetEditorController } from './use-editor.js';

type EditorProps = { editor: SpreadsheetEditorController };
type Renaming = { revision: string; label: string } | null;

export function PropertiesPanel({ editor: e }: { editor: SpreadsheetEditorController }) {
  const [titleDraft, setTitleDraft] = useState<string | null>(null);
  const title = titleDraft ?? e.workbook.name ?? 'Untitled spreadsheet';
  return (
    <div className="se-properties">
      <h3>Workbook properties</h3>
      <TextField
        label="Workbook title"
        value={title}
        onChange={setTitleDraft}
        isReadOnly={!e.canEdit}
      />
      <Button
        isDisabled={!e.canEdit || !title.trim()}
        onPress={() =>
          e.attempt(() => {
            e.commit({ ...e.workbook, name: title.trim() }, 'Rename workbook');
            setTitleDraft(null);
          })
        }
      >
        Update title
      </Button>
      <dl>
        <dt>Worksheet</dt>
        <dd>{e.sheet.name}</dd>
        <dt>Rows</dt>
        <dd>{e.sheet.data.length}</dd>
        <dt>Columns</dt>
        <dd>{Math.max(0, ...e.sheet.data.map((row) => row.length))}</dd>
        <dt>Selection</dt>
        <dd>{selectionAddress(e.selection)}</dd>
        <dt>Calculations</dt>
        <dd>{e.props.calculation?.mode ?? 'client'}</dd>
      </dl>
    </div>
  );
}

export function HistoryPanel({ editor: e }: { editor: SpreadsheetEditorController }) {
  if (e.props.versions && e.props.workbookId !== undefined) return <VersionsPanel editor={e} />;
  return (
    <div className="se-history">
      <h3>Editing history</h3>
      <p>Recent changes in this editing session.</p>
      {!e.history.past.length && <p>No changes in this session.</p>}
      <ol>
        {[...e.history.past].reverse().map((entry, index) => (
          <li key={index}>{entry.label}</li>
        ))}
      </ol>
      <Button isDisabled={!e.canWrite || !e.canUndo} onPress={e.undo}>
        Undo last change
      </Button>
      <Button isDisabled={!e.canWrite || !e.canRedo} onPress={e.redo}>
        Redo last change
      </Button>
    </div>
  );
}

/** Saves the workbook as a new named version. */
function SaveVersionForm({ editor: e }: EditorProps) {
  const [label, setLabel] = useState('');
  return (
    <form
      className="se-history__save"
      onSubmit={(event) => {
        event.preventDefault();
        void e.saveVersion(label.trim() || null).then((saved) => saved && setLabel(''));
      }}
    >
      <TextField label="Version name" value={label} onChange={setLabel} />
      <Button type="submit" size="sm">
        Save version
      </Button>
    </form>
  );
}

function RenameVersionForm({
  editor: e,
  version,
  renaming,
  setRenaming,
}: EditorProps & {
  version: WorkbookVersion;
  renaming: NonNullable<Renaming>;
  setRenaming(next: Renaming): void;
}) {
  return (
    <form
      className="se-history__rename"
      onSubmit={(event) => {
        event.preventDefault();
        void e
          .renameVersion(version, renaming.label.trim() || null)
          .then((renamed) => renamed && setRenaming(null));
      }}
    >
      <TextField
        label="Version name"
        value={renaming.label}
        onChange={(value) => setRenaming({ revision: version.revision, label: value })}
        autoFocus
      />
      <Button type="submit" size="sm">
        Rename
      </Button>
      <Button size="sm" variant="ghost" onPress={() => setRenaming(null)}>
        Cancel
      </Button>
    </form>
  );
}

/** A saved version, with its view, rename and restore actions in a context menu. */
function VersionEntry({
  editor: e,
  version,
  setRenaming,
}: EditorProps & { version: WorkbookVersion; setRenaming(next: Renaming): void }) {
  return (
    <ContextMenuTrigger>
      <Button variant="ghost" content="parts">
        <Attribution
          person={{ name: version.author ? e.personName(version.author) : 'Unknown' }}
          time={version.savedAt}
        />
        <span>{version.label ?? 'Saved automatically'}</span>
      </Button>
      <Menu
        aria-label="Version actions"
        items={[
          {
            id: 'open',
            label: 'View this version',
            ...editorCommandMetadata('open-history'),
            onAction: () => void e.openVersion(version),
          },
          ...(e.canRenameVersion
            ? [
                {
                  id: 'rename',
                  label: 'Rename version',
                  onAction: () =>
                    setRenaming({ revision: version.revision, label: version.label ?? '' }),
                },
              ]
            : []),
          ...(e.canRestoreVersion
            ? [
                {
                  id: 'restore',
                  label: 'Restore this version',
                  onAction: () => void e.restoreVersion(version),
                },
              ]
            : []),
        ]}
      />
    </ContextMenuTrigger>
  );
}

/** The workbook's saved versions, from the host's version store (SE-ADAPTERS-5.1). */
function VersionsPanel({ editor: e }: EditorProps) {
  const { versions, loading, error, refresh } = e.versionList;
  const [renaming, setRenaming] = useState<Renaming>(null);
  return (
    <div className="se-history" aria-label="Saved workbook history">
      <h3>Version history</h3>
      {e.canSaveVersion && <SaveVersionForm editor={e} />}
      <Button size="sm" onPress={refresh}>
        Refresh history
      </Button>
      {loading && <Spinner label="Loading history" />}
      {error && <InlineError>{error}</InlineError>}
      {!loading && !error && !versions.length && <p>No saved versions.</p>}
      <ol>
        {versions.map((version) => (
          <li
            key={version.revision}
            data-current={e.viewedVersion?.revision === version.revision || undefined}
          >
            {renaming?.revision === version.revision ? (
              <RenameVersionForm
                editor={e}
                version={version}
                renaming={renaming}
                setRenaming={setRenaming}
              />
            ) : (
              <VersionEntry editor={e} version={version} setRenaming={setRenaming} />
            )}
          </li>
        ))}
      </ol>
    </div>
  );
}
