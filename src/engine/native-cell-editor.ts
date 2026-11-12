/** Keep the vendor input inside the existing cell, including merged and frozen spans.
 * An in-flow input plus the cell's authored padding otherwise enlarges table rows.
 */
export function fitNativeCellEditor(
  cell: HTMLElement,
  control: HTMLInputElement | HTMLTextAreaElement,
): () => void {
  const style = getComputedStyle(cell);
  const position = cell.style.position;
  const positioned = style.position === 'static' || style.position === '';
  if (positioned) cell.style.position = 'relative';
  Object.assign(control.style, {
    position: 'absolute',
    top: '0px',
    left: '0px',
    width: '100%',
    height: '100%',
    minHeight: '0px',
    boxSizing: 'border-box',
    margin: '0px',
    padding: style.padding,
    font: style.font,
    color: style.color,
    backgroundColor: 'inherit',
    // Raw entry and its formula mirror share a stable starting edge, even in
    // a centered or right-aligned merged display cell.
    textAlign: 'left',
  });
  return () => {
    if (positioned) cell.style.position = position;
  };
}
