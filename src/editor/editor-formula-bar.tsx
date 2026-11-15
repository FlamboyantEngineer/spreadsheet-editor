'use client';

import { CodeEditor, type CodeTokenizer } from '@flamboyantengineer/ui-components/code-editor';
import type { CompletionSource } from '@flamboyantengineer/ui-components/completion-list';
import { FormulaGlyph } from '@flamboyantengineer/ui-components/core';
import { Icon } from '@flamboyantengineer/ui-components/icon';
import { Cluster, Stack } from '@flamboyantengineer/ui-components/layout';
import { TextField } from '@flamboyantengineer/ui-components/text-field';
import { VisuallyHidden } from '@flamboyantengineer/ui-components/visually-hidden';
import {
  type FocusEvent,
  type KeyboardEvent,
  type Ref,
  useId,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import { type CellSelection, selectionAddress } from '../model/editing.js';
import {
  buildFormulaReferenceTokenSegments,
  parseFormulaReferenceHighlights,
} from '../model/formula-reference-parser.js';
import { parseSpreadsheetRange } from '../model/spreadsheet-addressing.js';
import { formulaFunctionCatalog, type FormulaFunctionDescriptor } from './public-host.js';

/** Semantic tokens let the host use the same reference palette in the editor and grid. */
const tokenize: CodeTokenizer = (source) =>
  buildFormulaReferenceTokenSegments(source, parseFormulaReferenceHighlights(source)).map(
    ({ text, colorIndex }) => ({
      text,
      type: colorIndex === undefined ? undefined : `reference-${colorIndex}`,
    }),
  );

export interface EditorFormulaBarHandle {
  /** Inserts a picked grid reference over the current selection, without committing. */
  insertReference(range: string): void;
  focus(): void;
}

export interface EditorFormulaBarProps {
  formulaFunctions?: readonly FormulaFunctionDescriptor[];
  includeBuiltinFormulaFunctions?: boolean;
  ref?: Ref<EditorFormulaBarHandle>;
  selection: CellSelection;
  /** The host owns the protected draft, including updates while the same cell is edited. */
  draftValue: string;
  result: string;
  isReadOnly?: boolean;
  pending?: boolean;
  error?: boolean;
  draftDirty?: boolean;
  onDraftChange(value: string): void;
  /** The host commits and moves in a single transaction, avoiding a draft/navigation race. */
  onCommit(move: 'down' | 'stay'): void;
  onCancel(): void;
  onNavigate(selection: CellSelection): void;
  onEditingChange?(editing: boolean): void;
  /** Called only when focus leaves the bar and the host's editing scope. */
  onLeaveEditor?(): void;
  /** Includes host controls, portals and pointer reference picking in the editing scope. */
  keepsEditing?(target: EventTarget | null): boolean;
}

/** Address parsing must respect the workbook's selectable bounds, including during live input. */
function addressSelection(value: string): CellSelection | null {
  const range = parseSpreadsheetRange(value);
  if (!range || range.end.x > 16_383 || range.end.y > 1_048_575) return null;
  return { x1: range.start.x, y1: range.start.y, x2: range.end.x, y2: range.end.y };
}

/** Caller-owned drafts and transactions with the base library's formula layout and code surface. */
export function EditorFormulaBar({ ref, ...props }: EditorFormulaBarProps) {
  const completeFormula = useMemo<CompletionSource>(() => {
    const catalog = formulaFunctionCatalog(
      props.formulaFunctions,
      props.includeBuiltinFormulaFunctions,
    );
    return ({ value, query, isExplicit }) => {
      if (!value.trimStart().startsWith('=')) return [];
      if (!query && value.trim() !== '=' && !isExplicit) return [];
      const search = query.toUpperCase();
      return catalog
        .flatMap((descriptor) =>
          [descriptor.name, ...(descriptor.aliases ?? [])].map((name) => ({
            label: name,
            description: descriptor.description,
            insertText: descriptor.insertText ?? `${name}(`,
          })),
        )
        .filter((entry) => !search || entry.label.includes(search))
        .sort(
          (a, b) =>
            Number(b.label === search) - Number(a.label === search) ||
            a.label.localeCompare(b.label),
        );
    };
  }, [props.formulaFunctions, props.includeBuiltinFormulaFunctions]);
  const root = useRef<HTMLDivElement>(null);
  const input = useRef<HTMLTextAreaElement>(null);
  const editing = useRef(false);
  const composing = useRef(false);
  const pendingCaret = useRef<{ value: string; offset: number } | null>(null);
  const [addressDraft, setAddressDraft] = useState<string | null>(null);
  const [addressInvalid, setAddressInvalid] = useState(false);
  const address = selectionAddress(props.selection);
  const hintId = useId();
  const resultId = useId();
  const result = props.draftDirty
    ? 'Editing…'
    : props.pending
      ? 'Calculating…'
      : props.error
        ? 'Unavailable'
        : props.result;

  const setEditing = (value: boolean) => {
    if (editing.current === value) return;
    editing.current = value;
    props.onEditingChange?.(value);
  };

  useImperativeHandle(ref, () => ({
    insertReference(range) {
      if (props.isReadOnly || !range || !props.draftValue.trimStart().startsWith('=')) return;
      const start = input.current?.selectionStart ?? props.draftValue.length;
      const end = input.current?.selectionEnd ?? start;
      const value = props.draftValue.slice(0, start) + range + props.draftValue.slice(end);
      pendingCaret.current = { value, offset: start + range.length };
      props.onDraftChange(value);
      // A controlled parent may publish the draft later. Keep focus now and restore its caret
      // after that exact value arrives, rather than introducing an independent local draft.
      input.current?.focus();
      if (value === props.draftValue) {
        pendingCaret.current = null;
        input.current?.setSelectionRange(start + range.length, start + range.length);
      }
    },
    focus() {
      input.current?.focus();
    },
  }));

  useLayoutEffect(() => {
    const caret = pendingCaret.current;
    if (!caret || props.draftValue !== caret.value) return;
    pendingCaret.current = null;
    input.current?.focus();
    input.current?.setSelectionRange(caret.offset, caret.offset);
  }, [props.draftValue]);

  const isComposing = (event: KeyboardEvent) =>
    composing.current || event.nativeEvent.isComposing || event.nativeEvent.keyCode === 229;

  const onFormulaKeyDown = (event: KeyboardEvent) => {
    if (isComposing(event)) return;
    if (event.key !== 'Enter' && event.key !== 'Escape') return;
    event.preventDefault();
    event.stopPropagation();
    if (props.isReadOnly) return;
    // CodeEditor handles its completion list before this callback. For a single-line surface
    // it also prevents the unhandled Enter's line break before calling us, so defaultPrevented
    // cannot distinguish completion acceptance here.
    if (event.key === 'Enter') props.onCommit(event.ctrlKey || event.metaKey ? 'stay' : 'down');
    else props.onCancel();
    setEditing(false);
  };

  const onBlurCapture = (event: FocusEvent<HTMLDivElement>) => {
    const target = event.relatedTarget;
    if (target instanceof Node && root.current?.contains(target)) return;
    // A public ARIA relationship identifies this surface's portalled completion list.
    const listId = input.current?.getAttribute('aria-controls');
    const list = listId ? input.current?.ownerDocument.getElementById(listId) : null;
    if (target instanceof Node && list?.contains(target)) return;
    setEditing(false);
    if (props.keepsEditing?.(target)) return;
    props.onLeaveEditor?.();
  };

  const normalizeAddress = () => {
    setAddressDraft(null);
    setAddressInvalid(false);
  };

  return (
    <Cluster
      ref={root}
      wrap="nowrap"
      space="md"
      className="se-editor__formula-bar"
      onBlurCapture={onBlurCapture}
      onCompositionStartCapture={() => {
        composing.current = true;
      }}
      onCompositionEndCapture={() => {
        composing.current = false;
      }}
    >
      <Cluster
        className="se-editor__formula-widget"
        grow
        fill
        role="group"
        aria-label="Formula bar"
        wrap="nowrap"
        space="sm"
      >
        <TextField
          className="se-editor__address"
          label="Selected cell or range"
          density="compact"
          labelLayout="control-only"
          tracksChanges={false}
          size="sm"
          value={addressDraft ?? address}
          errorMessage={addressInvalid ? 'Enter a cell or range from A1 to XFD1048576.' : undefined}
          messageDisplay="tooltip"
          onFocus={() => setEditing(false)}
          onChange={(value) => {
            setAddressDraft(value);
            const selection = addressSelection(value);
            setAddressInvalid(!selection);
            if (selection) props.onNavigate(selection);
          }}
          onBlur={normalizeAddress}
          onKeyDown={(event) => {
            if (isComposing(event)) return;
            if (event.key === 'Escape') {
              event.preventDefault();
              event.stopPropagation();
              normalizeAddress();
            } else if (event.key === 'Enter') {
              event.preventDefault();
              event.stopPropagation();
              const selection = addressSelection(addressDraft ?? address);
              if (!selection) {
                setAddressInvalid(true);
                return;
              }
              props.onNavigate(selection);
              setAddressDraft(selectionAddress(selection));
              setAddressInvalid(false);
            }
          }}
        />
        <Icon icon={FormulaGlyph} />
        <Stack grow space="none">
          <CodeEditor
            ref={input}
            className="se-editor__formula-input"
            aria-label="Cell value or formula"
            aria-describedby={`${hintId} ${resultId}`}
            variant="surface"
            multiline={false}
            size="sm"
            value={props.draftValue}
            onChange={
              props.isReadOnly
                ? undefined
                : (value) => {
                    setEditing(true);
                    props.onDraftChange(value);
                  }
            }
            tokenize={tokenize}
            completions={completeFormula}
            isReadOnly={props.isReadOnly}
            onFocus={() => setEditing(!props.isReadOnly)}
            onKeyDown={onFormulaKeyDown}
          />
        </Stack>
      </Cluster>
      <output id={resultId} aria-label="Calculated cell value" title="Calculated value">
        {result}
      </output>
      <VisuallyHidden as="span" id={hintId}>
        {props.isReadOnly
          ? 'Cell content is read-only.'
          : 'Enter commits and moves down. Control/Command+Enter commits without moving. Escape cancels the edit.'}
      </VisuallyHidden>
    </Cluster>
  );
}
