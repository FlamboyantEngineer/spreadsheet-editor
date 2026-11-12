import { parseFragment, type DefaultTreeAdapterMap } from 'parse5';

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
/** parse5 decodes HTML entities without executing markup; work and output remain explicitly bounded. */
export function capturedText(value: string, policy: 'plain' | 'html-to-text'): string {
  if (value.length > 65_536) throw new RangeError('Captured text exceeds the HTML input limit.');
  let output = value;
  if (policy === 'html-to-text') {
    const root = parseFragment(value);
    const stack: { node: DefaultTreeAdapterMap['node']; depth: number }[] = [
      { node: root, depth: 0 },
    ];
    const parts: string[] = [];
    let visited = 0;
    while (stack.length) {
      const { node, depth } = stack.pop()!;
      if (++visited > 8192 || depth > 128)
        throw new RangeError('Captured HTML exceeds its tree limit.');
      if (node.nodeName === '#text') parts.push((node as DefaultTreeAdapterMap['textNode']).value);
      else if ('tagName' in node && node.tagName === 'br') parts.push('\n');
      else if ('childNodes' in node && !('tagName' in node && OMIT.has(node.tagName)))
        for (let i = node.childNodes.length - 1; i >= 0; i--)
          stack.push({ node: node.childNodes[i], depth: depth + 1 });
    }
    output = parts.join('');
  }
  // XML/Excel cannot preserve these control characters; never silently alter an operand.
  if (
    [...output].some((character) => {
      const code = character.codePointAt(0)!;
      return (
        (code < 32 && ![9, 10, 13].includes(code)) ||
        code === 0xfffe ||
        code === 0xffff ||
        (code >= 0xd800 && code <= 0xdfff)
      );
    })
  )
    throw new TypeError('Captured text contains characters Excel cannot represent.');
  if (output.length > 255)
    throw new RangeError('Excel formula strings have a 255-character limit.');
  return output;
}
