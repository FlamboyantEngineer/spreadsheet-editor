// This repository's settings for the checks shared through ui-components' tooling (its
// specs/TOOLING.md), which `npm run check:references`, `check:specs` and `check:structure` run.

export default {
  specs: { prefix: 'SE', hookPrefix: 'se' },
  structure: { names: ['src'] },
};
