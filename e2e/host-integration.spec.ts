import { expect, test, type Page, type Locator } from '@playwright/test';

const editor = (page: Page) => page.getByRole('region', { name: 'Host spreadsheet', exact: true });
const cell = (root: Locator, x: number, y = 0) =>
  root.locator(`td[data-x="${x}"][data-y="${y}"]`).first();
const formula = (root: Locator) =>
  root.getByRole('textbox', { name: 'Cell value or formula', exact: true });
test.beforeEach(async ({ page }) => {
  await page.goto('/host.html');
  await expect(cell(editor(page), 1)).toHaveText('6');
  await expect(cell(editor(page), 2)).toHaveText('42');
});

test('SE-HOST-1: protected formula drafts paint immediately, invalidate dependents and commit exactly once', async ({
  page,
}) => {
  const root = editor(page);
  await formula(root).fill('5');
  await expect(cell(root, 0)).toHaveText('5');
  await expect(cell(root, 1)).toHaveText('…');
  await expect(cell(root, 2)).toHaveText('42');
  await expect(page.getByLabel('Authored A1', { exact: true })).toHaveText('2');
  await expect(page.getByLabel('Host transaction count')).toHaveText('0');
  await expect(page.getByLabel('Host save count')).toHaveText('0');
  await formula(root).press('Escape');
  await expect(cell(root, 0)).toHaveText('2');
  await expect(cell(root, 1)).toHaveText('6');
  await formula(root).fill('5');
  await formula(root).press('Control+Enter');
  await expect(page.getByLabel('Authored A1', { exact: true })).toHaveText('5');
  await expect(page.getByLabel('Host transaction count')).toHaveText('1');
  await expect(cell(root, 1)).toHaveText('15');
  await expect(root.getByRole('textbox', { name: 'Selected cell or range' })).toHaveValue('A1');
  await root.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(page.getByLabel('Authored A1', { exact: true })).toHaveText('2');
  await expect(cell(root, 1)).toHaveText('6');
  await root.getByRole('button', { name: 'Redo', exact: true }).click();
  await expect(page.getByLabel('Authored A1', { exact: true })).toHaveText('5');
  await expect(cell(root, 1)).toHaveText('15');
});

test('SE-HOST-1: native drafts preserve accepted inputs, support cancellation and commit once', async ({
  page,
}) => {
  const root = editor(page);
  await cell(root, 0).dblclick();
  let native = cell(root, 0).getByRole('textbox');
  await native.fill('8');
  await expect(cell(root, 1)).toHaveText('…');
  await expect(cell(root, 2)).toHaveText('42');
  await expect(page.getByLabel('Authored A1', { exact: true })).toHaveText('2');
  await native.press('Escape');
  await expect(cell(root, 0)).toHaveText('2');
  await expect(cell(root, 1)).toHaveText('6');
  await cell(root, 0).dblclick();
  native = cell(root, 0).getByRole('textbox');
  await native.fill('8');
  await native.press('Enter');
  await expect(page.getByLabel('Authored A1', { exact: true })).toHaveText('8');
  await expect(cell(root, 1)).toHaveText('24');
  await expect(page.getByLabel('Host transaction count')).toHaveText('1');
});

test('SE-HOST-1: both formula surfaces advertise host functions, bare equals and IME completion priority', async ({
  page,
}) => {
  const root = editor(page);
  await formula(root).fill('=');
  await expect(page.getByRole('option').first()).toBeVisible();
  await formula(root).fill('=CAP');
  await expect(page.getByRole('option', { name: /CAPTURED/ }).first()).toBeVisible();
  await formula(root).dispatchEvent('compositionstart');
  await formula(root).dispatchEvent('keydown', { key: 'Enter', code: 'Enter', isComposing: true });
  await expect(formula(root)).toHaveValue('=CAP');
  await expect(page.getByLabel('Host transaction count')).toHaveText('0');
  await formula(root).dispatchEvent('compositionend');
  await page
    .getByRole('option', { name: /CAPTURED/ })
    .first()
    .click();
  await expect(formula(root)).toHaveValue('=CAPTURED(');
  await formula(root).fill('=CAPTURED()');
  await formula(root).press('Escape');
  await formula(root).press('Escape');
  await expect(formula(root)).toHaveValue('2');
  await cell(root, 0).dblclick();
  const native = cell(root, 0).getByRole('textbox');
  await native.fill('=');
  await expect(page.getByRole('option').first()).toBeVisible();
  await native.fill('=SNAPSHOTVALUE');
  await expect(page.getByRole('option', { name: /CAPTURED/ }).first()).toBeVisible();
  await native.dispatchEvent('keydown', {
    key: 'Enter',
    code: 'Enter',
    keyCode: 13,
    which: 13,
    isComposing: true,
  });
  await expect(native).toHaveValue('=SNAPSHOTVALUE');
  await native.press('Escape'); // close completions first
  await native.press('Escape'); // then cancel the draft
  await expect(page.getByLabel('Host transaction count')).toHaveText('0');
});

test('SE-HOST-1: external results update without edits, rich content and pins clear when pending, and result-only replacement preserves formulas', async ({
  page,
}) => {
  const root = editor(page);
  await page.getByRole('button', { name: 'Use authoritative results', exact: true }).click();
  await expect(cell(root, 3)).toContainText('Captured source');
  await expect(cell(root, 3).getByTestId('captured-rich-content')).toHaveText('Captured source');
  await expect(cell(root, 3)).toHaveAttribute('title', /Captured source revision/);
  await expect(cell(root, 3, 1)).toHaveText('100 cm');
  await root.getByRole('tab', { name: 'Another sheet', exact: true }).click();
  await expect(root.getByTestId('captured-rich-content')).toHaveCount(0);
  await expect(cell(root, 0)).toHaveText('7');
  await expect(cell(root, 0)).not.toHaveAttribute('title');
  await root.getByRole('tab', { name: 'Inputs', exact: true }).click();
  await expect(cell(root, 3).getByTestId('captured-rich-content')).toHaveText('Captured source');
  await page.getByRole('button', { name: 'Publish updated result', exact: true }).click();
  await expect(cell(root, 3)).toHaveText('Refreshed source');
  await expect(page.getByLabel('Host transaction count')).toHaveText('0');
  await root.getByRole('button', { name: 'Find', exact: true }).click();
  const search = root.getByRole('searchbox', { name: 'Find in workbook', exact: true });
  await search.fill('100 cm');
  await search.press('Enter');
  await expect(formula(root)).toHaveValue('=1 m');
  await search.fill('Refreshed source');
  await search.press('Enter');
  await expect(formula(root)).toHaveValue('=CAPTURED()');
  await root.getByRole('textbox', { name: 'Replace with', exact: true }).fill('Do not write');
  await expect(root.getByRole('button', { name: 'Replace', exact: true })).toBeDisabled();
  await expect(root.getByRole('button', { name: 'Replace all', exact: true })).toBeDisabled();
  await expect(formula(root)).toHaveValue('=CAPTURED()');
  await expect(page.getByLabel('Host transaction count')).toHaveText('0');
  await page.getByRole('button', { name: 'Publish pending result', exact: true }).click();
  await expect(cell(root, 3)).toHaveText('…');
  await expect(cell(root, 3).getByTestId('captured-rich-content')).toHaveCount(0);
  await expect(cell(root, 3)).not.toHaveAttribute('title');
  await expect(cell(root, 3)).not.toHaveAttribute('data-se-decoration');
  await expect(root.locator('.ui-find-bar').getByRole('status')).toHaveText('No results');
  await page.getByRole('button', { name: 'Publish escaped text', exact: true }).click();
  await expect(cell(root, 3)).toHaveText('<img src=x onerror=alert(1)> escaped source');
  await expect(cell(root, 3).locator('img,script')).toHaveCount(0);
  await page.getByRole('button', { name: 'Enable read-only', exact: true }).click();
  await expect(formula(root)).toHaveAttribute('readonly', '');
  await cell(root, 0).dblclick();
  await expect(cell(root, 0).getByRole('textbox')).toHaveCount(0);
});

test('SE-HOST-1: sheet switches restore selection, scroll and focus and commit the active draft', async ({
  page,
}) => {
  const root = editor(page);
  const address = root.getByRole('textbox', { name: 'Selected cell or range', exact: true });
  await address.fill('T80');
  await address.press('Enter');
  const viewport = root.locator('.jss_content').first();
  const scroll = await viewport.evaluate((element) => ({
    left: element.scrollLeft,
    top: element.scrollTop,
  }));
  expect(scroll.left).toBeGreaterThan(0);
  expect(scroll.top).toBeGreaterThan(0);
  await formula(root).fill('Draft committed on switch');
  await root.getByRole('tab', { name: 'Another sheet', exact: true }).click();
  await expect(cell(root, 1)).toHaveText('14');
  await root.getByRole('tab', { name: 'Inputs', exact: true }).click();
  await expect(address).toHaveValue('T80');
  await expect(cell(root, 19, 79)).toHaveText('Draft committed on switch');
  await expect(page.getByLabel('Host transaction count')).toHaveText('1');
  const restored = await viewport.evaluate((element) => ({
    left: element.scrollLeft,
    top: element.scrollTop,
  }));
  expect(restored.left).toBeCloseTo(scroll.left, 0);
  expect(restored.top).toBeCloseTo(scroll.top, 0);
  expect(
    await root
      .locator('.se-editor__grid-surface')
      .evaluate((element) => element === document.activeElement),
  ).toBe(true);
});

test('SE-HOST-1: host saves include active drafts and flush uses durability without adding a manual save', async ({
  page,
}) => {
  const root = editor(page);
  await formula(root).fill('9');
  await page.getByRole('button', { name: 'Save host editor', exact: true }).click();
  await expect(page.getByLabel('Saved A1', { exact: true })).toHaveText('9');
  await expect(page.getByLabel('Host save count')).toHaveText('1');
  await cell(root, 0).dblclick();
  const native = cell(root, 0).getByRole('textbox');
  await native.fill('11');
  await native.press('Control+s');
  await expect(page.getByLabel('Saved A1', { exact: true })).toHaveText('11');
  await expect(page.getByLabel('Host save count')).toHaveText('2');
  await page.getByRole('button', { name: 'Flush host editor', exact: true }).click();
  await expect(page.getByLabel('Host receipt count')).toHaveText('1');
  await expect(page.getByLabel('Host save count')).toHaveText('2');
  await page.getByRole('button', { name: 'Reject durability receipts', exact: true }).click();
  await cell(root, 0).click();
  await formula(root).fill('10');
  await page.getByRole('button', { name: 'Flush host editor', exact: true }).click();
  await expect(
    page.getByRole('alert').filter({ hasText: 'Host durability receipt was rejected.' }).first(),
  ).toBeVisible();
  await expect(page.getByLabel('Authored A1', { exact: true })).toHaveText('10');
  await expect(page.getByLabel('Host dirty state')).toHaveText('Unsaved');
  await expect(page.getByLabel('Host save count')).toHaveText('2');
});

test('SE-HOST-1: annotation-only worksheet edges reach the mounted grid without an authored edit', async ({
  page,
}) => {
  const root = editor(page);
  await root.getByRole('tab', { name: 'Anchored extent', exact: true }).click();
  await expect(cell(root, 3, 4)).toBeAttached();
  await expect(cell(root, 3, 4)).toHaveAttribute('data-se-comments', '1');
  await expect(page.getByLabel('Host transaction count')).toHaveText('0');
});
