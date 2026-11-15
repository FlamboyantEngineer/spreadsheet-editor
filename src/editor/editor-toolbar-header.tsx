'use client';

import { AvatarGroup } from '@flamboyantengineer/ui-components/avatar-group';
import { Button } from '@flamboyantengineer/ui-components/button';
import { SaveStatus } from '@flamboyantengineer/ui-components/save-status';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { collaboratorColorIndex } from '../engine/collaborator-color.js';
import { editorGlyphs } from './editor-commands.js';
import { MenuButton, type ToolbarSectionProps } from './editor-toolbar-controls.js';

const PRESENCE_STATUS_LABELS = { active: 'Active', idle: 'Idle', away: 'Away' } as const;
const MODE_LABELS = { author: 'Authoring', review: 'Reviewing', read: 'Reading' } as const;

/** The editing mode: a select in the full chrome, an icon menu in the embedded one. */
function ModeControl({ e, menus, embedded }: ToolbarSectionProps) {
  if (embedded)
    return (
      <MenuButton
        label="Editing mode"
        hint={`Editing mode: ${MODE_LABELS[e.mode]}`}
        icon={editorGlyphs[e.mode]}
        iconOnly
        disabled={e.locked}
        items={menus.review
          .slice(0, 3)
          .map((item) => ({ ...item, isSelected: item.id === e.mode }))}
      />
    );
  return (
    <SelectField
      tracksChanges={false}
      menu={[]}
      density="compact"
      labelLayout="control-only"
      size="sm"
      label="Editing mode"
      value={e.mode}
      isDisabled={e.locked}
      options={[
        { id: 'author', label: 'Authoring', isDisabled: !e.capabilities.edit },
        {
          id: 'review',
          label: 'Reviewing',
          isDisabled: !e.capabilities.suggest || e.props.features?.suggestions === false,
        },
        {
          id: 'read',
          label: 'Reading',
          isDisabled: !e.capabilities.edit && e.capabilities.suggest,
        },
      ]}
      onChange={(value) => e.setMode(value as typeof e.mode)}
    />
  );
}

/** The session's save status and the save button. */
function SaveControls({ e, embedded, part }: ToolbarSectionProps & { part: 'status' | 'button' }) {
  if (!e.props.session) return null;
  if (part === 'status')
    return (
      <SaveStatus
        phase={e.savePhase}
        onRetry={embedded ? undefined : () => void e.save()}
        hideSavedAfter={3000}
        labels={{
          dirty: 'Modified.',
          saving: e.saveAutomatic ? 'Autosaving…' : 'Saving…',
          saved: e.saveAutomatic ? 'Autosaved.' : 'Saved.',
        }}
      />
    );
  return (
    <Button
      size="sm"
      icon={editorGlyphs.save}
      label="Save"
      tooltip={e.savePhase === 'error' ? 'Retry save' : 'Save'}
      shortcut="Ctrl+S"
      isDisabled={!e.canEdit || e.savePhase === 'saving'}
      isPending={e.savePhase === 'saving'}
      onPress={() => e.save()}
    />
  );
}

function Collaborators({ e }: ToolbarSectionProps) {
  if (!e.peers.length) return null;
  return (
    <AvatarGroup
      aria-label="Active collaborators"
      size="xs"
      layout="strip"
      maxVisible={4}
      overflow="menu"
      people={e.peers.map((person) => ({
        id: person.id,
        name: person.name,
        colorIndex: collaboratorColorIndex(person.id),
        description: PRESENCE_STATUS_LABELS[person.status],
      }))}
      onAction={e.followCollaborator}
    />
  );
}

/** The title, save state, collaborators, contributors toggle and editing mode. */
export function ToolbarHeader(props: ToolbarSectionProps) {
  const { e, embedded } = props;
  return (
    <div className="se-editor__header">
      {!embedded && (
        <h2 className="se-editor__title">
          {e.props.title ?? e.workbook.name ?? 'Untitled spreadsheet'}
        </h2>
      )}
      <SaveControls {...props} part="status" />
      <Collaborators {...props} />
      <Button
        size="sm"
        variant="ghost"
        isSelected={e.showContributors}
        icon={editorGlyphs.contributors}
        label="Contributors"
        tooltip="Contributors"
        onPress={() => e.setShowContributors(!e.showContributors)}
      />
      {e.props.features?.modeControl !== false && <ModeControl {...props} />}
      <SaveControls {...props} part="button" />
    </div>
  );
}
