'use client';

import { CompletionList } from '@flamboyantengineer/ui-components/completion-list';
import {
  type CSSProperties,
  type FormEvent,
  type KeyboardEvent,
  type RefObject,
  useId,
  useLayoutEffect,
  useRef,
  useState,
} from 'react';
import { formulaFunctionCatalog, type FormulaFunctionDescriptor } from '../editor/public-host.js';
import {
  buildFormulaReferenceTokenSegments,
  parseFormulaReferenceHighlights,
} from '../model/formula-reference-parser.js';
import { requireDefined } from '../model/require-defined.js';
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-types.js';
import { fitNativeCellEditor } from './native-cell-editor.js';

type FormulaInput = HTMLInputElement | HTMLTextAreaElement;
type References = ReturnType<typeof parseFormulaReferenceHighlights>;
/** Where the native formula editor reports its drafts and which functions it completes. */
interface NativeFormulaOptions {
  onDraft: (value: string | null) => void;
  root: RefObject<HTMLDivElement | null>;
  functions?: readonly FormulaFunctionDescriptor[] | undefined;
  includeBuiltins?: boolean | undefined;
}

/** The mirror's box and text metrics, copied from the native control. */
function mirrorStyle(control: FormulaInput, container: HTMLDivElement): CSSProperties {
  const style = getComputedStyle(control);
  const box = control.getBoundingClientRect(),
    base = container.getBoundingClientRect();
  return {
    position: 'absolute',
    pointerEvents: 'none',
    zIndex: 6,
    overflow: 'hidden',
    whiteSpace: 'pre',
    display: 'flex',
    alignItems: control instanceof HTMLInputElement ? 'center' : 'flex-start',
    left: box.left - base.left - container.clientLeft,
    top: box.top - base.top - container.clientTop,
    width: box.width,
    height: box.height,
    font: style.font,
    padding: style.padding,
    lineHeight: style.lineHeight,
    color: style.color,
    boxSizing: 'border-box',
  };
}

/** Hides the control's own text behind the mirror, following its scroll; returns the undo. */
function hideBehindMirror(control: FormulaInput, onScroll: (left: number) => void) {
  const previous = { color: control.style.color, caretColor: control.style.caretColor };
  control.style.caretColor = getComputedStyle(control).color;
  control.style.color = 'transparent';
  const moved = () => onScroll(control.scrollLeft);
  moved();
  control.addEventListener('scroll', moved);
  return () => {
    control.removeEventListener('scroll', moved);
    control.style.color = previous.color;
    control.style.caretColor = previous.caretColor;
  };
}

/** The function name being typed before the caret of a formula, or nothing. */
function completionWord(value: string, prefix: string): string {
  if (!value.startsWith('=')) return '';
  return /(?:^|[=+*/^&<>,(;\s-])([A-Za-z][A-Za-z0-9.]*)$/.exec(prefix)?.[1] ?? '';
}

const completionText = (item: FormulaFunctionDescriptor) => item.insertText ?? item.name + '(';

/** Completes function names in the native editor through a completion list. */
function useFormulaCompletion(
  input: RefObject<FormulaInput | null>,
  {
    onDraft,
    setValue,
    functions,
    includeBuiltins = true,
  }: Omit<NativeFormulaOptions, 'root'> & {
    setValue: (value: string) => void;
  },
) {
  const [query, setQuery] = useState('');
  const [index, setIndex] = useState(0);
  const [open, setOpen] = useState(false);
  const id = useId();
  const options = formulaFunctionCatalog(functions, includeBuiltins)
    .filter((item) => [item.name, ...(item.aliases ?? [])].some((name) => name.startsWith(query)))
    .sort((a, b) => Number(b.name === query) - Number(a.name === query))
    .slice(0, 12);
  function track() {
    const current = input.current;
    if (!current) return;
    onDraft(current.value);
    setValue(current.value);
    const prefix = current.value.slice(0, current.selectionStart ?? current.value.length);
    const word = completionWord(current.value, prefix);
    setQuery(word.toUpperCase());
    setIndex(0);
    setOpen(!!word || prefix === '=');
  }
  function insert(text: string, replaceWord = false) {
    const current = input.current;
    if (!current) return;
    const end = current.selectionEnd ?? current.value.length;
    const start = (current.selectionStart ?? end) - (replaceWord ? query.length : 0);
    current.setRangeText(text, Math.max(0, start), end, 'end');
    current.focus({ preventScroll: true });
    onDraft(current.value);
    setValue(current.value);
    setOpen(false);
  }
  function onKey(event: KeyboardEvent<HTMLElement>) {
    if (!input.current || event.nativeEvent.isComposing || !open || !options.length) return;
    if (event.key === 'ArrowDown') setIndex((index + 1) % options.length);
    else if (event.key === 'ArrowUp') setIndex((index + options.length - 1) % options.length);
    else if (event.key === 'Enter' || event.key === 'Tab')
      insert(completionText(options[index]), true);
    else if (event.key === 'Escape') setOpen(false);
    else return;
    event.preventDefault();
    event.stopPropagation();
  }
  const popup = (
    <CompletionList
      aria-label="Formula functions"
      id={id}
      anchorRef={input}
      isOpen={open && options.length > 0}
      onOpenChange={setOpen}
      items={options.map((item) => ({
        id: item.name,
        label: item.name,
        description: `${item.signature} — ${item.description}`,
      }))}
      activeKey={options[index]?.name ?? null}
      onActiveKeyChange={(key) => setIndex(options.findIndex((item) => item.name === key))}
      onAction={(key) => {
        const item = options.find((item) => item.name === key);
        if (item) insert(completionText(item), true);
      }}
    />
  );
  return { track, insert, onKey, popup, close: () => setOpen(false) };
}

/** The formula's text, each reference in its colour, scrolled with the control. */
function referenceOverlay(
  value: string,
  references: References,
  { mirror, scroll }: { mirror: CSSProperties; scroll: number },
) {
  return (
    <div className="se-editor__native-formula" aria-hidden="true" style={mirror}>
      <span style={{ display: 'block', flex: '0 0 auto', transform: `translateX(${-scroll}px)` }}>
        {buildFormulaReferenceTokenSegments(value, references).map((segment, index) => (
          <span
            key={index}
            style={{
              color:
                segment.colorIndex === undefined
                  ? undefined
                  : `var(--se-reference-${segment.colorIndex}, ${segment.color})`,
            }}
          >
            {segment.text}
          </span>
        ))}
      </span>
    </div>
  );
}

const formulaControl = (element: unknown): element is FormulaInput =>
  element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement;

/** React completion UI attached to the exact native editor published by CE's edition tuple. */
export function useNativeFormula({
  onDraft,
  root,
  functions,
  includeBuiltins,
}: NativeFormulaOptions) {
  const input = useRef<FormulaInput | null>(null);
  const releaseEditor = useRef<(() => void) | null>(null);
  const [value, setValue] = useState('');
  const references = value.startsWith('=') ? parseFormulaReferenceHighlights(value) : [];
  const highlighted = references.length > 0;
  const [mirror, setMirror] = useState<CSSProperties>({});
  const [scroll, setScroll] = useState(0);
  useLayoutEffect(() => {
    const control = input.current,
      container = root.current;
    if (!control || !container || !highlighted) return;
    setMirror(mirrorStyle(control, container));
    return hideBehindMirror(control, setScroll);
  }, [highlighted, value, root]);
  const completion = useFormulaCompletion(input, { onDraft, setValue, functions, includeBuiltins });
  function start(sheet: JSpreadsheetWorksheetInstance | null) {
    queueMicrotask(() => {
      const edition = (sheet as { edition?: [HTMLElement] } | null)?.edition;
      const control = edition?.[0]?.firstElementChild;
      if (!formulaControl(control)) return;
      releaseEditor.current?.();
      releaseEditor.current = fitNativeCellEditor(requireDefined(edition)[0], control);
      input.current = control;
      completion.track();
    });
  }
  function end() {
    releaseEditor.current?.();
    releaseEditor.current = null;
    input.current = null;
    setValue('');
    completion.close();
    onDraft(null);
  }
  function onInput(event: FormEvent<HTMLElement>) {
    if (!formulaControl(event.target)) return;
    input.current = event.target;
    completion.track();
  }
  const overlay = highlighted && referenceOverlay(value, references, { mirror, scroll });
  const { insert, onKey, popup } = completion;
  return { input, start, end, insert, onInput, onKey, popup, overlay };
}
