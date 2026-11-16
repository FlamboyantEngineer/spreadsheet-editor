'use client';

import { CodeEditor, type CodeTokenizer } from '@flamboyantengineer/ui-components/code-editor';
import type { CompletionSource } from '@flamboyantengineer/ui-components/completion-list';
import { FormulaGlyph } from '@flamboyantengineer/ui-components/core';
import { Icon } from '@flamboyantengineer/ui-components/icon';
import { Cluster, Stack } from '@flamboyantengineer/ui-components/layout';
import { VisuallyHidden } from '@flamboyantengineer/ui-components/visually-hidden';
import {
  type FocusEvent,
  type KeyboardEvent,
  type Ref,
  type RefObject,
  useId,
  useImperativeHandle,
  useLayoutEffect,
  useMemo,
  useRef,
} from 'react';
import type { CellSelection } from '../model/editing.js';
import {
  buildFormulaReferenceTokenSegments,
  parseFormulaReferenceHighlights,
} from '../model/formula-reference-parser.js';
import { AddressField, claimKey } from './editor-address-field.js';
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

/** Completes function names from the host catalog while a formula is typed. */
function formulaCompletions(catalog: ReturnType<typeof formulaFunctionCatalog>): CompletionSource {
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
          Number(b.label === search) - Number(a.label === search) || a.label.localeCompare(b.label),
      );
  };
}

const resultText = ({ draftDirty, pending, error, result }: EditorFormulaBarProps) =>
  draftDirty ? 'Editing…' : pending ? 'Calculating…' : error ? 'Unavailable' : result;

/** The bar's handle: reference insertion keeps the caller-owned draft and restores its caret. */
function useFormulaHandle(
  ref: Ref<EditorFormulaBarHandle> | undefined,
  input: RefObject<HTMLTextAreaElement | null>,
  { isReadOnly, draftValue, onDraftChange }: EditorFormulaBarProps,
) {
  const pendingCaret = useRef<{ value: string; offset: number } | null>(null);
  useImperativeHandle(ref, () => ({
    insertReference(range) {
      if (isReadOnly || !range || !draftValue.trimStart().startsWith('=')) return;
      const start = input.current?.selectionStart ?? draftValue.length;
      const end = input.current?.selectionEnd ?? start;
      const value = draftValue.slice(0, start) + range + draftValue.slice(end);
      pendingCaret.current = { value, offset: start + range.length };
      onDraftChange(value);
      // A controlled parent may publish the draft later. Keep focus now and restore its caret
      // after that exact value arrives, rather than introducing an independent local draft.
      input.current?.focus();
      if (value === draftValue) {
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
    if (!caret || draftValue !== caret.value) return;
    pendingCaret.current = null;
    input.current?.focus();
    input.current?.setSelectionRange(caret.offset, caret.offset);
  }, [draftValue, input]);
}

/** Editing state, composition tracking, and the bar's Enter, Escape and focus-leave handling. */
function useFormulaEditing(
  props: EditorFormulaBarProps,
  {
    root,
    input,
  }: { root: RefObject<HTMLDivElement | null>; input: RefObject<HTMLTextAreaElement | null> },
) {
  const editing = useRef(false);
  const composing = useRef(false);
  const setEditing = (value: boolean) => {
    if (editing.current === value) return;
    editing.current = value;
    props.onEditingChange?.(value);
  };
  const isComposing = (event: KeyboardEvent) =>
    composing.current || event.nativeEvent.isComposing || event.nativeEvent.keyCode === 229;

  const onFormulaKeyDown = (event: KeyboardEvent) => {
    if (isComposing(event)) return;
    if (event.key !== 'Enter' && event.key !== 'Escape') return;
    claimKey(event);
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
  const composition = {
    onCompositionStartCapture: () => {
      composing.current = true;
    },
    onCompositionEndCapture: () => {
      composing.current = false;
    },
  };
  return { setEditing, isComposing, onFormulaKeyDown, onBlurCapture, composition };
}

/** Caller-owned drafts and transactions with the base library's formula layout and code surface. */
export function EditorFormulaBar({ ref, ...props }: EditorFormulaBarProps) {
  const { formulaFunctions, includeBuiltinFormulaFunctions } = props;
  const completeFormula = useMemo(
    () =>
      formulaCompletions(formulaFunctionCatalog(formulaFunctions, includeBuiltinFormulaFunctions)),
    [formulaFunctions, includeBuiltinFormulaFunctions],
  );
  const root = useRef<HTMLDivElement>(null);
  const input = useRef<HTMLTextAreaElement>(null);
  const hintId = useId();
  const resultId = useId();
  useFormulaHandle(ref, input, props);
  const { setEditing, isComposing, onFormulaKeyDown, onBlurCapture, composition } =
    useFormulaEditing(props, { root, input });

  return (
    <Cluster
      ref={root}
      wrap="nowrap"
      space="md"
      className="se-editor__formula-bar"
      onBlurCapture={onBlurCapture}
      {...composition}
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
        <AddressField
          selection={props.selection}
          isComposing={isComposing}
          onFocus={() => setEditing(false)}
          onNavigate={props.onNavigate}
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
        {resultText(props)}
      </output>
      <VisuallyHidden as="span" id={hintId}>
        {props.isReadOnly
          ? 'Cell content is read-only.'
          : 'Enter commits and moves down. Control/Command+Enter commits without moving. Escape cancels the edit.'}
      </VisuallyHidden>
    </Cluster>
  );
}
