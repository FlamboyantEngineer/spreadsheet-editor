/** Returns a value the caller knows is present, and throws when it is not. */
export function requireDefined<T>(value: T | null | undefined, what = 'a value'): T {
  if (value === undefined || value === null) throw new Error(`Expected ${what} to be present.`);
  return value;
}
