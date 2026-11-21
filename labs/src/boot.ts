import { bootLabSettings } from '@flamboyantengineer/ui-components/labs';
import { labStorage } from './settings.js';

// Preserve links from the earlier lab picker using the shared setting's current spelling.
const url = new URL(location.href);
if (url.searchParams.get('stylesheet') === 'unstyled') {
  url.searchParams.set('stylesheet', 'none');
  history.replaceState(null, '', url);
}
bootLabSettings({
  storage: labStorage(),
  referenceStylesheet: '/reference.css',
  shadcnStylesheet: '/shadcn.css',
});
