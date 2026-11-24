import { expect, type Locator, type Page, test } from '@playwright/test';

async function openEditor(page: Page) {
  await page.goto('/full.html');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await expect(editor.getByRole('group', { name: 'Workbook grid', exact: true })).toBeVisible({
    timeout: 10_000,
  });
  return editor;
}

const address = (editor: Locator) =>
  editor.getByRole('textbox', { name: 'Selected cell or range', exact: true });
const formula = (editor: Locator) =>
  editor.getByRole('textbox', { name: 'Cell value or formula', exact: true });
const cell = (editor: Locator, x: number, y: number) =>
  editor.locator(`td[data-x="${x}"][data-y="${y}"]`).first();

test('SE-CHROME-3.1: the grid scopes keyboard input, restores navigation, and lets Tab leave', async ({
  page,
}) => {
  const editor = await openEditor(page);
  const grid = editor.getByRole('group', { name: 'Workbook grid', exact: true });
  const b2 = cell(editor, 1, 1);
  await b2.click();
  await page.getByRole('button', { name: 'Calculate on server', exact: true }).focus();
  await page.keyboard.press('Delete');
  await expect(b2).toHaveText('1e-25');

  await grid.focus();
  await page.keyboard.press('ArrowRight');
  await expect(address(editor)).toHaveValue('C2');
  await page.keyboard.press('Shift+F10');
  await expect(page.getByRole('menu', { name: 'Spreadsheet context menu' })).toBeVisible();
  await page.keyboard.press('Escape');
  await page.keyboard.press('ArrowRight');
  await expect(address(editor)).toHaveValue('D2');
  await page.keyboard.press('Tab');
  await expect(grid.locator(':focus')).toHaveCount(0);
  await expect(grid).not.toBeFocused();
  await expect(address(editor)).toHaveValue('D2');
});

test('SE-CHROME-3.1: the Find shortcut moves focus to its query and Escape returns it', async ({
  page,
}) => {
  const editor = await openEditor(page);
  await cell(editor, 1, 1).click();
  await page.keyboard.press('Control+f');
  const find = editor.getByRole('searchbox', { name: 'Find in workbook', exact: true });
  await expect(find).toBeFocused();
  await page.keyboard.type('Near one');
  await expect(find).toHaveValue('Near one');
  await expect(cell(editor, 1, 1)).toHaveText('1e-25');
  await page.keyboard.press('Escape');
  await expect(editor.getByRole('group', { name: 'Workbook grid', exact: true })).toBeFocused();
});

test('SE-CHROME-4.1: leaving a row header makes Delete clear a cell instead of deleting its row', async ({
  page,
}) => {
  const editor = await openEditor(page);
  await editor.locator('tr[data-y="1"] .jss_row').first().click();
  await expect(address(editor)).toHaveValue(/^A2:[A-Z]+2$/);
  await page.keyboard.press('ArrowDown');
  await expect(address(editor)).toHaveValue('A3');
  await page.keyboard.press('Delete');
  await expect(cell(editor, 0, 2)).toHaveText('');
  await expect(cell(editor, 1, 2)).toHaveText('0.9999999999999999999999999');
  await expect(cell(editor, 0, 3)).toHaveText('Recover difference');
});

test('SE-CHROME-4.1: Go To refuses an address outside the mounted worksheet', async ({ page }) => {
  const editor = await openEditor(page);
  await cell(editor, 1, 1).click();
  await address(editor).fill('AA101');
  await address(editor).press('Enter');
  await expect(address(editor)).toHaveAttribute('aria-invalid', 'true');
  await expect(formula(editor)).toHaveValue('1E-25');
  await formula(editor).fill('44');
  await formula(editor).press('Control+Enter');
  await expect(cell(editor, 1, 1)).toHaveText('44');
  await expect(address(editor)).toHaveValue('B2');
});

test('SE-CHROME-3.1: native clipboard shortcuts retain formulas without granting read permission', async ({
  page,
}) => {
  const editor = await openEditor(page);
  const source = cell(editor, 0, 0);
  await source.click();
  await formula(editor).fill('=1+1');
  await formula(editor).press('Control+Enter');
  await expect(source).toHaveText('2');
  await source.click();
  await page.keyboard.press('Control+x');
  await cell(editor, 1, 0).click();
  await page.keyboard.press('Control+v');
  await expect(source).toHaveText('');
  await expect(formula(editor)).toHaveValue('=1+1');
});

test('SE-CHROME-3.1: menu and keyboard clipboard paths share structured copies', async ({
  page,
  context,
}) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  const editor = await openEditor(page);
  const source = cell(editor, 0, 0);
  await source.click();
  await formula(editor).fill('=1+1');
  await formula(editor).press('Control+Enter');
  await expect(source).toHaveText('2');

  await source.click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Cut', exact: true }).click();
  await expect.poll(() => page.evaluate(() => navigator.clipboard.readText())).toBe('2');
  await cell(editor, 1, 0).click();
  await page.keyboard.press('Control+v');
  await expect(source).toHaveText('');
  await expect(formula(editor)).toHaveValue('=1+1');

  await page.keyboard.press('Control+c');
  await cell(editor, 2, 0).click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Paste', exact: true }).click();
  await expect(formula(editor)).toHaveValue('=1+1');
  await expect(cell(editor, 2, 0)).toHaveText('2');
});

test('SE-CHROME-3.1: malformed native clipboard content goes through the editor notice', async ({
  page,
}) => {
  const editor = await openEditor(page);
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  const b2 = cell(editor, 1, 1);
  await b2.click();
  await b2.evaluate((target) => {
    const transfer = new DataTransfer();
    transfer.setData('text/plain', '"unterminated');
    target.dispatchEvent(
      new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: transfer }),
    );
  });
  await expect(editor.getByRole('alert')).toContainText('Unclosed quoted field');
  await expect(b2).toHaveText('1e-25');
  expect(errors).toEqual([]);
});
