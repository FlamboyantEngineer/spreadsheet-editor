import { type DefaultTreeAdapterMap, parseFragment } from 'parse5';
import { requireDefined } from '../model/require-defined.js';

const OMIT = new Set([
  'script',
  'style',
  'iframe',
  'object',
  'embed',
  'template',
  'noscript',
  'svg',
  'math',
  'canvas',
  'head',
  'title',
  'textarea',
  'select',
  'input',
  'button',
]);
type Node = DefaultTreeAdapterMap['node'];

/** The text of an HTML fragment, with line breaks for `br` and without inert or form content. */
function htmlText(value: string): string {
  const stack: { node: Node; depth: number }[] = [{ node: parseFragment(value), depth: 0 }];
  const parts: string[] = [];
  let visited = 0;
  while (stack.length) {
    const { node, depth } = requireDefined(stack.pop());
    if (++visited > 8192 || depth > 128)
      throw new RangeError('Captured HTML exceeds its tree limit.');
    if (node.nodeName === '#text') parts.push((node as DefaultTreeAdapterMap['textNode']).value);
    else if ('tagName' in node && node.tagName === 'br') parts.push('\n');
    else if ('childNodes' in node && !('tagName' in node && OMIT.has(node.tagName)))
      stack.push(
        ...[...node.childNodes].reverse().map((child) => ({ node: child, depth: depth + 1 })),
      );
  }
  return parts.join('');
}

/** Control characters, noncharacters and lone surrogates, which XML cannot carry. */
function unrepresentable(character: string): boolean {
  const code = requireDefined(character.codePointAt(0));
  return (
    (code < 32 && ![9, 10, 13].includes(code)) ||
    code === 0xfffe ||
    code === 0xffff ||
    (code >= 0xd800 && code <= 0xdfff)
  );
}

/** parse5 decodes HTML entities without executing markup; work and output remain explicitly bounded. */
export function capturedText(
  value: string,
  policy: 'plain' | 'html-to-text',
  maxLength: 255 | 32767 = 255,
): string {
  if (value.length > 65_536) throw new RangeError('Captured text exceeds the HTML input limit.');
  const output = policy === 'html-to-text' ? htmlText(value) : value;
  // XML/Excel cannot preserve these control characters; never silently alter an operand.
  if ([...output].some(unrepresentable))
    throw new TypeError('Captured text contains characters Excel cannot represent.');
  if (output.length > maxLength)
    throw new RangeError(
      maxLength === 255
        ? 'Excel formula strings have a 255-character limit.'
        : 'Excel cell strings have a 32767-character limit.',
    );
  return output;
}
