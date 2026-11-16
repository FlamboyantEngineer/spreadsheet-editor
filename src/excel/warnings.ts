import type { XlsxWarning } from './contracts.js';

/** One finding of an import or export, counted under its code. */
export interface ReportedWarning {
  code: string;
  message: string;
  location: string;
  disposition?: XlsxWarning['disposition'];
}

/** Warnings counted by code, with up to five distinct locations each. */
export function createWarnings(disposition: XlsxWarning['disposition']) {
  const warnings = new Map<string, XlsxWarning>();
  return {
    warn(warning: ReportedWarning) {
      const { code, message, location } = warning;
      const entry = warnings.get(code) ?? {
        code,
        message,
        disposition: warning.disposition ?? disposition,
        count: 0,
        locations: [],
      };
      entry.count++;
      if (entry.locations.length < 5 && !entry.locations.includes(location))
        entry.locations.push(location);
      warnings.set(code, entry);
    },
    list: () => [...warnings.values()],
  };
}
