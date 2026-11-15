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
import type { JSpreadsheetWorksheetInstance } from './jspreadsheet-host.js';
import { fitNativeCellEditor } from './native-cell-editor.js';

/** React completion UI attached to the exact native editor published by CE's edition tuple. */
export function useNativeFormula(
  onDraft: (value: string | null) => void,
  root: RefObject<HTMLDivElement | null>,
  functions?: readonly FormulaFunctionDescriptor[],
  includeBuiltins = true,
) {
  const input = useRef<HTMLInputElement | HTMLTextAreaElement | null>(null);
  const releaseEditor = useRef<(() => void) | null>(null);
  const [value, setValue] = useState('');
  const [mirror, setMirror] = useState<CSSProperties>({});
  const [scroll, setScroll] = useState(0);
  const references = value.startsWith('=') ? parseFormulaReferenceHighlights(value) : [];
  const highlighted = references.length > 0;
  useLayoutEffect(() => {
    const control = input.current,
      container = root.current;
    if (!control || !container || !highlighted) return;
    const style = getComputedStyle(control);
    const box = control.getBoundingClientRect(),
      base = container.getBoundingClientRect();
    const previous = { color: control.style.color, caretColor: control.style.caretColor };
    setMirror({
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
    });
    control.style.caretColor = style.color;
    control.style.color = 'transparent';
    const moved = () => setScroll(control.scrollLeft);
    moved();
    control.addEventListener('scroll', moved);
    return () => {
      control.removeEventListener('scroll', moved);
      control.style.color = previous.color;
      control.style.caretColor = previous.caretColor;
    };
  }, [highlighted, value, root]);
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
    const word = current.value.startsWith('=')
      ? (/(?:^|[=+*/^&<>,(;\s-])([A-Za-z][A-Za-z0-9.]*)$/.exec(prefix)?.[1] ?? '')
      : '';
    setQuery(word.toUpperCase());
    setIndex(0);
    setOpen(!!word || prefix === '=');
  }
  function start(sheet: JSpreadsheetWorksheetInstance | null) {
    queueMicrotask(() => {
      const edition = (sheet as { edition?: [HTMLElement] } | null)?.edition;
      const control = edition?.[0]?.firstElementChild;
      if (control instanceof HTMLInputElement || control instanceof HTMLTextAreaElement) {
        releaseEditor.current?.();
        releaseEditor.current = fitNativeCellEditor(requireDefined(edition)[0], control);
        input.current = control;
        track();
      }
    });
  }
  function end() {
    releaseEditor.current?.();
    releaseEditor.current = null;
    input.current = null;
    setValue('');
    setOpen(false);
    onDraft(null);
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
  function onInput(event: FormEvent<HTMLElement>) {
    if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement) {
      input.current = event.target;
      track();
    }
  }
  function onKey(event: KeyboardEvent<HTMLElement>) {
    if (!input.current || event.nativeEvent.isComposing || !open || !options.length) return;
    if (event.key === 'ArrowDown') setIndex((index + 1) % options.length);
    else if (event.key === 'ArrowUp') setIndex((index + options.length - 1) % options.length);
    else if (event.key === 'Enter' || event.key === 'Tab')
      insert(options[index].insertText ?? options[index].name + '(', true);
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
        if (item) insert(item.insertText ?? item.name + '(', true);
      }}
    />
  );
  const overlay = highlighted && (
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
  return { input, start, end, insert, onInput, onKey, popup, overlay };
}
