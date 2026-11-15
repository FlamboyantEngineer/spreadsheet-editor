import { expect, type Locator, type Page, test } from '@playwright/test';
import { requireDefined } from '../src/model/require-defined.js';

test.use({ viewport: { width: 1440, height: 1000 } });

const ROWS = 2_000;

async function openLarge(page: Page) {
  await page.goto(`/large.html?rows=${ROWS}`);
  const editor = page.locator('.se-editor').first();
  await expect(editor.locator('td[data-x="24"][data-y="3"]')).toHaveText('24', {
    timeout: 30_000,
  });
  return editor;
}

/** The cell under a point of the grid viewport, as `y:text`. */
function cellAt(editor: Locator, from: 'top' | 'bottom') {
  return editor.locator('.jss_content').evaluate((content, side) => {
    const bounds = content.getBoundingClientRect();
    const y = side === 'top' ? bounds.top + 40 : bounds.top + content.clientHeight - 12;
    const row = document.elementFromPoint(bounds.left + 250, y)?.closest('tr');
    return `${row?.dataset.y}:${row?.querySelector('td[data-x="3"]')?.textContent}`;
  }, from);
}

function attachedRows(editor: Locator) {
  return editor.locator('.jss_worksheet tbody tr[data-y]').count();
}

test('SE-PERF-1: a 50,000-cell worksheet attaches only the rows near the viewport', async ({
  page,
}) => {
  const editor = await openLarge(page);
  const viewport = editor.locator('.jss_content');
  expect(await attachedRows(editor)).toBeLessThan(100);
  // The scrollbar spans every row at its own height.
  const heights = await viewport.evaluate((content) => ({
    scroll: content.scrollHeight,
    header: content.querySelector('thead')?.getBoundingClientRect().height ?? Number.NaN,
  }));
  expect(heights.scroll - heights.header).toBeCloseTo(ROWS * 24, -1);
  await viewport.evaluate((content) => {
    content.scrollTop = content.scrollHeight;
  });
  await expect.poll(() => cellAt(editor, 'bottom')).toBe(`${ROWS - 1}:Item ${ROWS}.4`);
  expect(await attachedRows(editor)).toBeLessThan(100);
  // Halfway down, the row under the viewport's top is the one its offset names.
  const expected = await viewport.evaluate((content) => {
    content.scrollTop = 24 * 1_000;
    const header = content.querySelector('thead')?.getBoundingClientRect().height ?? Number.NaN;
    return Math.floor((content.scrollTop + 40 - header) / 24);
  });
  await expect.poll(() => cellAt(editor, 'top')).toBe(`${expected}:Item ${expected + 1}.4`);
});

test('SE-PERF-1: keyboard moves past the attached rows scroll to the selected cell', async ({
  page,
}) => {
  const editor = await openLarge(page);
  await editor.locator('td[data-x="3"][data-y="2"]').click();
  await page.keyboard.press('Control+ArrowDown');
  const last = editor.locator(`td[data-x="3"][data-y="${ROWS - 1}"]`);
  await expect(last).toHaveClass(/highlight-selected/);
  await expect(last).toBeInViewport();
  await page.keyboard.press('Control+ArrowUp');
  const first = editor.locator('td[data-x="3"][data-y="0"]');
  await expect(first).toHaveClass(/highlight-selected/);
  await expect(first).toBeInViewport();
});

test('SE-PERF-1: rows attached after a zoom change carry the zoomed text metrics', async ({
  page,
}) => {
  const editor = await openLarge(page);
  const fontSize = (cell: Locator) =>
    cell.evaluate((element) => getComputedStyle(element).fontSize);
  const near = editor.locator('td[data-x="3"][data-y="2"]');
  const base = Number.parseFloat(await fontSize(near));
  await editor
    .getByRole('group', { name: 'Spreadsheet zoom', exact: true })
    .getByRole('button', { name: 'Zoom in', exact: true })
    .click();
  await expect.poll(async () => Number.parseFloat(await fontSize(near))).toBeGreaterThan(base);
  const zoomed = await fontSize(near);
  await editor.locator('.jss_content').evaluate((content) => {
    content.scrollTop = content.scrollHeight;
  });
  const far = editor.locator(`td[data-x="3"][data-y="${ROWS - 1}"]`);
  await expect(far).toBeInViewport();
  expect(await fontSize(far)).toBe(zoomed);
});

test('SE-PERF-1: a far cell can be edited and recalculates its row', async ({ page }) => {
  const editor = await openLarge(page);
  await editor.locator('.jss_content').evaluate((content) => {
    content.scrollTop = content.scrollHeight;
  });
  const y = ROWS - 2;
  const cell = editor.locator(`td[data-x="1"][data-y="${y}"]`);
  await expect(cell).toBeInViewport();
  await cell.dblclick();
  await cell.getByRole('textbox').fill('1');
  await page.keyboard.press('Enter');
  // The total sums A:C of its row: (y + 1) + 1 + 3(y + 1).
  await expect(editor.locator(`td[data-x="24"][data-y="${y}"]`)).toHaveText(
    String(4 * (y + 1) + 1),
  );
  await expect(editor.locator(`td[data-x="1"][data-y="${y + 1}"]`)).toBeInViewport();
});

test('SE-PERF-1: find reveals a match in a detached row', async ({ page }) => {
  const editor = await openLarge(page);
  await editor.locator('td[data-x="3"][data-y="2"]').click();
  await page.keyboard.press('Control+f');
  const search = editor.getByRole('searchbox', { name: 'Find in workbook', exact: true });
  await search.fill('Item 1800.4');
  await search.press('Enter');
  const match = editor.locator('td[data-x="3"][data-y="1799"]');
  await expect(match).toHaveClass(/highlight-selected/);
  await expect(match).toBeInViewport();
});

test('SE-PERF-1: a selection reaching detached rows keeps its outline on the attached ones', async ({
  page,
}) => {
  const editor = await openLarge(page);
  await editor.locator('td[data-x="3"][data-y="2"]').click();
  await editor.locator('.jss_content').evaluate((content) => {
    content.scrollTop = content.scrollHeight;
  });
  const last = editor.locator(`td[data-x="5"][data-y="${ROWS - 1}"]`);
  await expect(last).toBeInViewport();
  await last.click({ modifiers: ['Shift'] });
  await expect(last).toHaveClass(/highlight/);
  const outline = editor.locator('.se-editor__selection-outline');
  await expect(outline).toHaveCount(1);
  const [box, cell] = await Promise.all([outline.boundingBox(), last.boundingBox()]);
  // The outline ends at the last row and runs off the top of the viewport.
  expect(
    Math.abs(
      requireDefined(box).y +
        requireDefined(box).height -
        (requireDefined(cell).y + requireDefined(cell).height),
    ),
  ).toBeLessThan(2);
  expect(await outline.evaluate((element) => getComputedStyle(element).borderTopWidth)).toBe('0px');
});
