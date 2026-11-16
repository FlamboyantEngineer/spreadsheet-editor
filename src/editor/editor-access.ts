import {
  type EditorCapabilities,
  FULL_CAPABILITIES,
  READ_ONLY_CAPABILITIES,
  resolveSpreadsheetMode,
} from '../host/capabilities.js';
import type { SpreadsheetEditorProps, SpreadsheetMode } from './contracts.js';

/** The mode the editor works in: the requested one where the capabilities allow it. */
function effectiveMode(
  props: SpreadsheetEditorProps,
  capabilities: EditorCapabilities,
  { historical, requestedMode }: { historical: boolean; requestedMode: SpreadsheetMode },
): SpreadsheetMode {
  const highestMode = resolveSpreadsheetMode(capabilities, { historical });
  if (highestMode === 'author') return requestedMode;
  const reviewable =
    (capabilities.suggest && props.features?.suggestions !== false) ||
    (capabilities.comment && props.features?.comments !== false);
  return highestMode === 'review' && reviewable ? 'review' : 'read';
}

/** What the reader may do: a viewed version is read-only whatever the host allows. */
export function editorAccess(
  props: SpreadsheetEditorProps,
  view: { historical: boolean; requestedMode: SpreadsheetMode },
) {
  const hostCapabilities = props.capabilities ?? FULL_CAPABILITIES;
  const capabilities = view.historical ? READ_ONLY_CAPABILITIES : hostCapabilities;
  /** Nothing is allowed: no mode can be chosen. */
  const locked = resolveSpreadsheetMode(capabilities, { historical: view.historical }) === 'read';
  const mode = effectiveMode(props, capabilities, view);
  const suggestions = props.features?.suggestions !== false;
  const canEdit = mode === 'author' && capabilities.edit;
  const canSuggest = mode === 'review' && capabilities.suggest && suggestions;
  return {
    hostCapabilities,
    capabilities,
    locked,
    mode,
    canEdit,
    canSuggest,
    canWrite: canEdit || canSuggest,
    canComment: mode !== 'read' && capabilities.comment && props.features?.comments !== false,
    canDecide: canEdit && capabilities.decide && suggestions,
  };
}
