'use client';

import { Attribution } from '@flamboyantengineer/ui-components/attribution';
import { Button } from '@flamboyantengineer/ui-components/button';
import { ChatBubbleCheckGlyph, ChatBubbleXMarkGlyph } from '@flamboyantengineer/ui-components/core';
import { SelectField } from '@flamboyantengineer/ui-components/select-field';
import { useState } from 'react';
import { requireDefined } from '../model/require-defined.js';
import type { SpreadsheetSuggestion, SuggestedCellChange } from '../model/review.js';
import { parseSpreadsheetRange } from '../model/spreadsheet-addressing.js';
import type { SpreadsheetEditorController } from './use-editor.js';

type EditorProps = { editor: SpreadsheetEditorController };

/** Accept and reject buttons for a pending suggestion. */
export function SuggestionDecisionButtons({
  editor: e,
  suggestion,
}: EditorProps & { suggestion: SpreadsheetSuggestion }) {
  return (
    <div className="se-suggestion__actions">
      <Button
        startIcon={ChatBubbleCheckGlyph}
        isDisabled={!e.canDecide || suggestion.outdated}
        onPress={() => e.attempt(() => e.decide(suggestion.id, 'accepted'))}
      >
        Accept suggestion
      </Button>
      <Button
        startIcon={ChatBubbleXMarkGlyph}
        isDisabled={!e.canDecide}
        onPress={() => e.attempt(() => e.decide(suggestion.id, 'rejected'))}
      >
        Reject suggestion
      </Button>
    </div>
  );
}

/** One changed cell: a button that selects it, and its before and after values. */
function SuggestionChange({
  editor: e,
  suggestion,
  change,
}: EditorProps & { suggestion: SpreadsheetSuggestion; change: SuggestedCellChange }) {
  return (
    <div className="se-suggestion__change">
      <Button
        onPress={() => {
          const cell = requireDefined(parseSpreadsheetRange(change.address));
          e.select(
            { x1: cell.start.x, x2: cell.end.x, y1: cell.start.y, y2: cell.end.y },
            change.sheetId,
          );
        }}
        isDisabled={suggestion.outdated}
      >
        {e.workbook.worksheets.find((sheet) => sheet.id === change.sheetId)?.name}!{change.address}
      </Button>
      {change.before !== change.after ? (
        <p>
          <del>{String(change.before ?? '(empty)')}</del>
          <span aria-hidden="true"> → </span>
          <ins>{String(change.after ?? '(empty)')}</ins>
        </p>
      ) : (
        <p>Cell formatting change</p>
      )}
    </div>
  );
}

function SuggestionCard({
  editor: e,
  suggestion,
}: EditorProps & { suggestion: SpreadsheetSuggestion }) {
  return (
    <article
      className="se-suggestion"
      aria-label={`Suggestion: ${suggestion.summary}`}
      data-status={suggestion.status}
    >
      <h3>{suggestion.summary}</h3>
      <Attribution person={{ name: suggestion.author.name }} time={suggestion.createdAt} />
      <p className="se-suggestion__status">
        {suggestion.status}
        {suggestion.outdated ? ' · deleted cells' : ''}
      </p>
      {suggestion.structure && <p>Worksheet structure change, including any affected formulas.</p>}
      {suggestion.changes.map((change) => (
        <SuggestionChange
          key={`${change.sheetId}:${change.address}`}
          editor={e}
          suggestion={suggestion}
          change={change}
        />
      ))}
      {suggestion.commentThreadId && (
        <Button size="sm" onPress={() => e.openSuggestion(suggestion.id)}>
          Open discussion
        </Button>
      )}
      {suggestion.status === 'pending' && (
        <SuggestionDecisionButtons editor={e} suggestion={suggestion} />
      )}
      {suggestion.decidedBy && (
        <Attribution
          person={{ name: suggestion.decidedBy.name }}
          time={suggestion.decidedAt}
          template="Decided {time} by {person}"
        />
      )}
    </article>
  );
}

export function SuggestionsPanel({ editor: e }: EditorProps) {
  const [filter, setFilter] = useState('pending');
  const suggestions = (e.workbook.suggestions ?? []).filter(
    (item) => filter === 'all' || item.status === filter,
  );
  return (
    <div className="se-suggestions" aria-label="Workbook suggestions">
      <SelectField
        label="Suggestion filter"
        value={filter}
        onChange={(value) => setFilter(String(value))}
        options={['pending', 'accepted', 'rejected', 'all'].map((id) => ({
          id,
          label: id[0].toUpperCase() + id.slice(1),
        }))}
      />
      <p>
        In Reviewing mode, cell edits and formatting become suggestions. Canonical values change
        after acceptance.
      </p>
      {!suggestions.length && <p>No {filter === 'all' ? '' : filter + ' '}suggestions.</p>}
      {suggestions.map((suggestion) => (
        <SuggestionCard key={suggestion.id} editor={e} suggestion={suggestion} />
      ))}
    </div>
  );
}
