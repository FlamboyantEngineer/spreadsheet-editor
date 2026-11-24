import { SaxesParser } from 'saxes';

interface XmlEvents {
  open?(name: string, attributes: Readonly<Record<string, string>>, depth: number): void;
  close?(depth: number): void;
  text?(value: string): void;
}

const decoder = new TextDecoder('utf-8', { fatal: true });

/** Reads one bounded OOXML part with matching open and close events, including empty elements. */
export function readXlsxXml(bytes: Uint8Array, events: XmlEvents = {}): void {
  const parser = new SaxesParser({ xmlns: false });
  let depth = 0;
  parser.on('doctype', () => {
    throw new TypeError('XML declarations for document types and entities are unsupported.');
  });
  parser.on('opentag', ({ name, attributes }) => {
    if (++depth > 128) throw new RangeError('XLSX XML nesting exceeds its depth limit.');
    events.open?.(name.slice(name.lastIndexOf(':') + 1), attributes, depth);
  });
  parser.on('closetag', () => {
    events.close?.(depth);
    depth--;
  });
  const onText = (value: string) => events.text?.(value);
  parser.on('text', onText);
  parser.on('cdata', onText);
  parser.write(decoder.decode(bytes)).close();
}
