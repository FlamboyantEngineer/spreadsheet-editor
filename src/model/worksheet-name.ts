/** The form worksheet names are compared in: spaces collapsed, case folded. */
export function worksheetNameKey(name: string): string {
  return name.trim().replace(/\s+/g, ' ').toUpperCase();
}
