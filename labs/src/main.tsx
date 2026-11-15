import { useState, useSyncExternalStore } from 'react';
import { createRoot } from 'react-dom/client';
import { UiProvider } from '@flamboyantengineer/ui-components/core';
import { mountLabBar } from '@flamboyantengineer/ui-components/labs';
import { labStorage } from './settings.js';
import { Button } from '@flamboyantengineer/ui-components/button';
import { SpreadsheetEditor, createWorkbook } from '../../src/index.js';
import { createMemoryHost } from '../../src/host/index.js';

const alex = createMemoryHost({ reader: { id: 'alex', name: 'Alex' } }).identity;
import type { Workbook } from '../../src/model/index.js';
import { FullLab } from './full-lab.js';
import { HostLab } from './host-lab.js';
import { LargeLab } from './large-lab.js';
import { engineeringSample } from './sample.js';
import 'jsuites/dist/jsuites.css';
import 'jspreadsheet-ce/dist/jspreadsheet.css';
import './lab.css';

const labs = [
  {
    id: 'minimal',
    title: 'Minimal',
    description: 'An editable grid with worksheet tabs. No toolbar, menus or formula bar.',
    detail:
      'Double-click a cell to edit it. Right-click for cell actions, or use the keyboard to navigate and edit.',
  },
  {
    id: 'embedded',
    title: 'Embedded in an app',
    description: 'A compact editor inside a project workspace.',
    detail:
      'The application owns navigation and context. The editor supplies compact tools, formulas, review panels and save callbacks.',
  },
  {
    id: 'full',
    title: 'Full editor',
    description: 'The complete workspace: menus, formatting, formulas, comments and suggestions.',
    detail:
      'Explore the complete editor, then use the sections below to try server calculation, exact values and both document insertion modes.',
  },
  {
    id: 'host',
    title: 'Host integration',
    description:
      'Public lifecycle handles, captured results, function descriptors and React cell content.',
    detail:
      'Explore a host-controlled workbook with protected drafts, publication revisions, explicit save and durability callbacks.',
  },
  {
    id: 'large',
    title: 'Large worksheet',
    description: 'A 10,000-cell worksheet for checking how the grid responds to typing and edits.',
    detail:
      'Type into cells and commit edits on a worksheet with 10,000 cells and a formula in every row. The browser tests hold its typing and commit times to a budget.',
  },
] as const;
const current = labs.find((lab) => location.pathname === `/${lab.id}.html`);
let navigationSearch = '';
const navigationListeners = new Set<() => void>();
function subscribeNavigation(listener: () => void) {
  navigationListeners.add(listener);
  return () => navigationListeners.delete(listener);
}
mountLabBar({
  host: document.getElementById('lab-chrome')!,
  title: current?.title ?? 'Spreadsheet editor labs',
  indexHref: current ? '/' : null,
  storage: labStorage(),
  referenceStylesheet: '/reference.css',
  onSettingsChange(_settings, search) {
    navigationSearch = search;
    for (const listener of navigationListeners) listener();
  },
});

function MinimalLab() {
  return (
    <section className="lab-demo" aria-label="Minimal editor demo">
      <SpreadsheetEditor
        chrome="minimal"
        aria-label="Minimal spreadsheet"
        defaultValue={createWorkbook([
          ['Item', 'Quantity', 'Unit price', 'Total'],
          ['Apples', '4', '0.5', '=B2*C2'],
          ['Pears', '6', '0.75', '=B3*C3'],
          ['Total', '=SUM(B2:B3)', '', '=SUM(D2:D3)'],
        ])}
      />
    </section>
  );
}

function EmbeddedLab() {
  const [workbook, setWorkbook] = useState<Workbook>(() => structuredClone(engineeringSample));
  const [saved, setSaved] = useState(false);
  return (
    <section className="lab-app" aria-label="Example project workspace">
      <div className="lab-app__topbar">
        <strong>Engineering workspace</strong>
        <span>Example workspace</span>
        <span className="lab-app__user">AL</span>
      </div>
      <div className="lab-app__body">
        <aside className="lab-app__nav" aria-label="Project context">
          <span className="lab-app__eyebrow">PROJECT</span>
          <strong>Falcon X</strong>
          <span className="lab-app__branch">Engineering</span>
          <span className="lab-app__active">Mass budget</span>
          <div className="lab-app__nav-note">
            Spreadsheet lab
            <br />
            Local example data
          </div>
        </aside>
        <div className="lab-app__content">
          <p className="lab-app__breadcrumb">Falcon X / Engineering / Mass budget</p>
          <div className="lab-app__heading">
            <div>
              <span className="lab-app__eyebrow">ENGINEERING WORKBOOK</span>
              <h2>Payload mass budget</h2>
            </div>
            <span className="lab-app__revision">Revision A</span>
          </div>
          <p className="lab-app__description">
            Subsystem allocations, calculated totals and remaining margin.
          </p>
          <SpreadsheetEditor
            chrome="embedded"
            aria-label="Embedded spreadsheet"
            value={workbook}
            onChange={(next) => {
              setWorkbook(next);
              setSaved(false);
            }}
            title="Mass budget"
            identity={alex}
            onSave={async () => {
              setSaved(true);
            }}
          />
          <div className="lab-app__footnote" role="status">
            {saved
              ? 'Saved in this lab session.'
              : 'Edits stay in this lab session. Select Save to exercise the host callback.'}
          </div>
        </div>
      </div>
    </section>
  );
}

function App() {
  const search = useSyncExternalStore(subscribeNavigation, () => navigationSearch);
  const href = (id?: string) => `${id ? `/${id}.html` : '/'}${search}`;
  return (
    <UiProvider>
      <main className="lab-page">
        <header className="lab-page__intro">
          <h1>{current?.title ?? 'Spreadsheet editor labs'}</h1>
          <p>
            {current?.detail ??
              'Explore the same spreadsheet editor in standalone and host-controlled examples. Each lab is interactive and uses the shared ui-components controls and reference stylesheet.'}
          </p>
        </header>
        {current ? (
          <>
            <nav className="lab-switcher" aria-label="Editor labs">
              {labs.map((lab) => (
                <a
                  key={lab.id}
                  href={href(lab.id)}
                  aria-current={lab.id === current.id ? 'page' : undefined}
                >
                  {lab.title}
                </a>
              ))}
            </nav>
            {current.id === 'minimal' ? (
              <MinimalLab />
            ) : current.id === 'embedded' ? (
              <EmbeddedLab />
            ) : current.id === 'host' ? (
              <HostLab />
            ) : current.id === 'large' ? (
              <LargeLab />
            ) : (
              <FullLab />
            )}
          </>
        ) : (
          <section className="lab-index" aria-label="Editor examples">
            <h2>Editor configurations</h2>
            <ul>
              {labs.map((lab) => (
                <li key={lab.id}>
                  <a href={href(lab.id)}>{lab.title}</a>
                  <p>{lab.description}</p>
                  <code>
                    {lab.id === 'host'
                      ? 'editorRef + external results'
                      : lab.id === 'large'
                        ? 'chrome="embedded", 10,000 cells'
                        : `chrome="${lab.id}"`}
                  </code>
                </li>
              ))}
            </ul>
            <p className="lab-index__note">
              All examples support exact arithmetic and cell context menus. The full editor also
              demonstrates client/server calculation and document insertion.
            </p>
          </section>
        )}
        {current && (
          <footer className="lab-page__footer">
            <a href={href()}>← All spreadsheet labs</a>
            <Button
              size="sm"
              variant="ghost"
              onPress={() => window.scrollTo({ top: 0, behavior: 'smooth' })}
            >
              Back to top
            </Button>
          </footer>
        )}
      </main>
    </UiProvider>
  );
}

createRoot(document.getElementById('root')!).render(<App />);
