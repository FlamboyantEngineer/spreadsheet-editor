import { expect, test, type Page } from '@playwright/test';

test.use({ viewport: { width: 1440, height: 1000 } });

interface Probe {
  events: number[];
  longTasks: number[];
  mutations: Record<string, number>;
}

/**
 * What the keys pressed in `action` cost: the slowest event the browser reports
 * (input handling up to the next paint), the main-thread long tasks added
 * together, which also catch work an event schedules after it returns, and the
 * grid body's DOM mutations by kind.
 */
async function measure(page: Page, action: () => Promise<void>) {
  await page.evaluate(() => {
    const probe: Probe = { events: [], longTasks: [], mutations: {} };
    (window as unknown as { gridBudget: Probe }).gridBudget = probe;
    new PerformanceObserver((list) => {
      for (const entry of list.getEntries())
        if (/^key|input$/.test(entry.name)) probe.events.push(entry.duration);
    }).observe({ type: 'event', durationThreshold: 16 } as PerformanceObserverInit);
    new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) probe.longTasks.push(entry.duration);
    }).observe({ type: 'longtask' });
    new MutationObserver((records) => {
      for (const record of records) {
        const kind =
          `${record.type} ${record.target.nodeName} ${record.attributeName ?? ''}`.trim();
        probe.mutations[kind] = (probe.mutations[kind] ?? 0) + 1;
      }
    }).observe(document.querySelector('.se-editor .jss_worksheet tbody')!, {
      subtree: true,
      childList: true,
      attributes: true,
      characterData: true,
    });
  });
  await action();
  // Event entries are delivered after the next paint; let them arrive.
  await page.waitForTimeout(1_000);
  return page.evaluate(() => {
    const probe = (window as unknown as { gridBudget: Probe }).gridBudget;
    return {
      slowestEvent: Math.max(0, ...probe.events),
      longTasks: probe.longTasks.reduce((sum, duration) => sum + duration, 0),
      mutations: probe.mutations,
    };
  });
}

async function openLargeWorksheet(page: Page) {
  await page.goto('/large.html');
  const editor = page.locator('.se-editor').first();
  await expect(editor.locator('td[data-x="24"][data-y="399"]')).toHaveText('2400', {
    timeout: 20_000,
  });
  // 10,000 authored cells; the grid always shows at least 26 columns.
  await expect(editor.locator('.jss_worksheet tbody td[data-x]:not(.jss_row)')).toHaveCount(10_400);
  return editor;
}

// SE-PERF-1: 10,000 cells smooth, 50,000 usable. The mutation budgets are exact
// and are what hold the grid to work on changed cells only. Wall-clock time on a
// shared CI runner varies with what else runs there, so the time budgets only catch
// a gross regression and leave room for a busy runner.
test('SE-PERF-1: typing into a cell of a 10,000-cell worksheet rewrites no other cell', async ({
  page,
}) => {
  const editor = await openLargeWorksheet(page);
  const cell = editor.locator('td[data-x="1"][data-y="3"]');
  await cell.dblclick();
  await expect(cell.getByRole('textbox')).toBeFocused();
  await page.keyboard.press('End');
  await page.keyboard.press('1');
  const typing = await measure(page, async () => {
    for (const key of '2345') {
      await page.keyboard.press(key);
      await page.waitForTimeout(150);
    }
  });
  const report = JSON.stringify(typing);
  const mutations = Object.values(typing.mutations).reduce((sum, count) => sum + count, 0);
  expect(mutations, report).toBeLessThanOrEqual(20);
  expect(typing.slowestEvent, report).toBeLessThan(600);
  expect(typing.longTasks, report).toBeLessThan(2_000);
});

test('SE-PERF-1: committing a cell of a 10,000-cell worksheet keeps the grid rows', async ({
  page,
}) => {
  const editor = await openLargeWorksheet(page);
  const cell = editor.locator('td[data-x="1"][data-y="3"]');
  await cell.dblclick();
  await cell.getByRole('textbox').fill('12345');
  const commit = await measure(page, () => page.keyboard.press('Enter'));
  // Row 4's total sums A4:C4 = 4 + 12345 + 12.
  await expect(editor.locator('td[data-x="24"][data-y="3"]')).toHaveText('12361');
  const report = JSON.stringify(commit);
  // Rebuilding the rows would replace every row of the body.
  expect(commit.mutations['childList TBODY'] ?? 0, report).toBe(0);
  // Presentation is rewritten for changed cells only, not restyled for every cell.
  expect(commit.mutations['attributes TD style'] ?? 0, report).toBeLessThanOrEqual(20);
  expect(commit.slowestEvent, report).toBeLessThan(4_000);
  expect(commit.longTasks, report).toBeLessThan(4_000);
});
