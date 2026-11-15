import { expect, test, type Page } from '@playwright/test';

const source = (page: Page) =>
  page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
async function openLab(page: Page) {
  await page.goto('/full.html');
  await expect(source(page).locator('td[data-x="1"][data-y="1"]').first()).toHaveText('1e-25');
}
async function selectMode(page: Page, label: string) {
  await source(page)
    .getByRole('button', { name: /Editing mode$/ })
    .click();
  await page.getByRole('option', { name: label, exact: true }).click();
}
async function contextOn(page: Page, x: number, y: number) {
  await source(page).locator(`td[data-x="${x}"][data-y="${y}"]`).first().click({ button: 'right' });
  const menu = page.getByRole('menu', { name: 'Spreadsheet context menu', exact: true });
  await expect(menu).toBeVisible();
  return menu;
}

test('SE-CHROME-1: full chrome formats cells without rounding canonical calculations', async ({
  page,
}) => {
  await openLab(page);
  const editor = source(page);
  await expect(editor.getByRole('menubar', { name: 'Spreadsheet menu' })).toBeVisible();
  await expect(editor.getByRole('toolbar', { name: 'Spreadsheet formatting' })).toBeVisible();
  const nearOne = editor.locator('td[data-x="1"][data-y="2"]').first();
  await nearOne.click();
  await editor.getByRole('button', { name: 'Bold selection', exact: true }).click();
  await expect(nearOne).toHaveCSS('font-weight', '700');
  await editor.getByRole('button', { name: /Number format$/ }).click();
  await page.getByRole('option', { name: 'Percent', exact: true }).click();
  await expect(nearOne).toHaveText('100%');
  await expect(editor.locator('td[data-x="2"][data-y="2"]').first()).toHaveText('FALSE');
  await editor.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(nearOne).toHaveText('0.9999999999999999999999999');
  await expect(editor.getByLabel('Selection summary')).toContainText('Count: 1');
  await editor
    .getByRole('group', { name: 'Spreadsheet zoom', exact: true })
    .getByRole('button', { name: 'Zoom in', exact: true })
    .click();
  await expect(editor.getByRole('group', { name: 'Spreadsheet zoom', exact: true })).toContainText(
    '125%',
  );
});

test('SE-CONTEXT-1: mouse and keyboard menus use React and rewrite row references', async ({
  page,
}) => {
  await openLab(page);
  const menu = await contextOn(page, 1, 1);
  await expect(page.locator('.jss_contextmenu')).not.toBeVisible();
  await menu.getByRole('menuitem', { name: 'Insert rows above', exact: true }).click();
  const editor = source(page);
  await expect(editor.locator('td[data-x="1"][data-y="2"]').first()).toHaveText('1e-25');
  await expect(editor.locator('td[data-x="1"][data-y="3"]').first()).toHaveText(
    '0.9999999999999999999999999',
  );
  await editor.locator('td[data-x="1"][data-y="2"]').first().click();
  await page.keyboard.press('Shift+F10');
  await expect(
    page.getByRole('menu', { name: 'Spreadsheet context menu', exact: true }),
  ).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(
    page.getByRole('menu', { name: 'Spreadsheet context menu', exact: true }),
  ).not.toBeVisible();
});

test('SE-REVIEW-3: comments support replies, resolve/reopen, and host persistence', async ({
  page,
}) => {
  await openLab(page);
  const menu = await contextOn(page, 1, 3);
  await menu.getByRole('menuitem', { name: 'Add comment', exact: true }).click();
  const editor = source(page);
  const draft = page.getByRole('textbox', { name: 'Comment on Calculations!B4', exact: true });
  await draft.fill('Please double-check this exact result.');
  await page.getByRole('button', { name: 'Post comment', exact: true }).click();
  const thread = page.getByRole('group', { name: 'Comment thread Calculations!B4', exact: true });
  await thread
    .getByRole('textbox', { name: 'Reply to comment', exact: true })
    .fill('Checked without rounding.');
  await thread.getByRole('button', { name: 'Post reply', exact: true }).click();
  await expect(thread).toContainText('Checked without rounding.');
  await expect(editor.locator('td[data-x="1"][data-y="3"]').first()).toHaveAttribute(
    'data-se-comments',
    '1',
  );
  await thread.getByRole('button', { name: 'Resolve', exact: true }).click();
  await expect(thread).not.toBeVisible();
  await editor.getByRole('button', { name: /^Comments \(/ }).click();
  await editor.getByLabel('Comment filter', { exact: true }).click();
  await page.getByRole('option', { name: 'Resolved comments', exact: true }).click();
  await editor
    .getByRole('button')
    .filter({ hasText: 'Please double-check this exact result.' })
    .click();
  await thread.getByRole('button', { name: 'Reopen', exact: true }).click();
  await page.keyboard.press('Escape');
  await editor.getByLabel('Comment filter', { exact: true }).click();
  await page.getByRole('option', { name: 'Open comments', exact: true }).click();
  await editor.getByRole('button', { name: 'Save', exact: true }).click();
  await page.getByRole('button', { name: 'Reset workbook', exact: true }).click();
  await expect(
    editor.getByText('Please double-check this exact result.', { exact: true }),
  ).not.toBeVisible();
  await page.getByRole('button', { name: 'Reload saved workbook', exact: true }).click();
  await expect(
    editor.getByText('Please double-check this exact result.', { exact: true }),
  ).toBeVisible();
  await expect(editor.locator('td[data-x="1"][data-y="3"]').first()).toHaveText('1e-25');
});

test('SE-REVIEW-4: native grid edits create suggestions and accept/reject preserves canonical data', async ({
  page,
}) => {
  await openLab(page);
  await selectMode(page, 'Reviewing');
  const editor = source(page),
    cell = editor.locator('td[data-x="1"][data-y="1"]').first();
  await cell.dblclick();
  await cell.getByRole('textbox').fill('2E-25');
  await cell.getByRole('textbox').press('Enter');
  await expect(cell).toHaveText('2e-25');
  await selectMode(page, 'Authoring');
  await expect(cell).toHaveText('1e-25');
  const suggestion = editor.getByRole('article', { name: 'Suggestion: Edit cells', exact: true });
  await expect(suggestion).toContainText('2E-25');
  await suggestion.getByRole('button', { name: 'Accept suggestion', exact: true }).click();
  await expect(cell).toHaveText('2e-25');
  await expect(editor.locator('td[data-x="1"][data-y="3"]').first()).toHaveText('2e-25');
  await selectMode(page, 'Reviewing');
  await cell.click();
  await editor.getByRole('textbox', { name: 'Cell value or formula', exact: true }).fill('3E-25');
  await editor
    .getByRole('textbox', { name: 'Cell value or formula', exact: true })
    .press('Control+Enter');
  await selectMode(page, 'Authoring');
  const second = editor.getByRole('article', { name: 'Suggestion: Edit B2', exact: true });
  await second.getByRole('button', { name: 'Reject suggestion', exact: true }).click();
  await expect(cell).toHaveText('2e-25');
});

test('SE-CHROME-2: worksheet context actions rename and preserve cross-sheet formulas', async ({
  page,
}) => {
  await openLab(page);
  const editor = source(page);
  await editor.getByRole('tab', { name: 'Inputs', exact: true }).click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Rename', exact: true }).click();
  const name = editor.getByRole('textbox', { name: 'Rename Inputs', exact: true });
  await name.fill('Parameters');
  await name.press('Enter');
  await expect(editor.getByRole('tab', { name: 'Calculations', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(editor.locator('td[data-x="1"][data-y="6"]').first()).toHaveText('3e-25');
  await expect(editor.getByRole('tab', { name: 'Parameters', exact: true })).toBeVisible();
});

test('SE-REVIEW-5: host read-only mode locks edits, comments, decisions and context actions', async ({
  page,
}) => {
  await openLab(page);
  await page.getByRole('button', { name: 'Read-only host', exact: true }).click();
  const editor = source(page);
  await expect(
    editor.getByRole('textbox', { name: 'Cell value or formula', exact: true }),
  ).toHaveAttribute('readonly', '');
  await expect(editor.getByRole('button', { name: 'Bold selection', exact: true })).toBeDisabled();
  const menu = await contextOn(page, 1, 1);
  await expect(menu.getByRole('menuitem', { name: 'Clear contents', exact: true })).toHaveAttribute(
    'aria-disabled',
    'true',
  );
  await expect(menu.getByRole('menuitem', { name: 'Add comment', exact: true })).toHaveAttribute(
    'aria-disabled',
    'true',
  );
  await page.keyboard.press('Escape');
  await editor.getByRole('button', { name: /^Suggestions \(/ }).click();
  await expect(
    editor.getByRole('button', { name: 'Accept suggestion', exact: true }),
  ).toBeDisabled();
});

test('SE-CHROME-3: search fields receive keyboard input without editing the selected grid cell', async ({
  page,
}) => {
  await openLab(page);
  const editor = source(page);
  await editor.locator('td[data-x="1"][data-y="1"]').first().click();
  await editor.getByRole('button', { name: 'Find', exact: true }).click();
  await editor
    .getByRole('searchbox', { name: 'Find in workbook', exact: true })
    .fill('Small number');
  await editor.getByRole('searchbox', { name: 'Find in workbook', exact: true }).press('Enter');
  await expect(
    editor.getByRole('textbox', { name: 'Cell value or formula', exact: true }),
  ).toHaveValue('Small number');
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveText('1e-25');
  await editor
    .getByRole('textbox', { name: 'Replace with', exact: true })
    .fill('Tiny exact number');
  await editor.getByRole('button', { name: 'Replace', exact: true }).click();
  await expect(editor.locator('td[data-x="0"][data-y="1"]').first()).toHaveText(
    'Tiny exact number',
  );
});

test('SE-REVIEW-6: bulk clipboard paste becomes one proposal and never edits canonical cells directly', async ({
  page,
  context,
}) => {
  await context.grantPermissions(['clipboard-read', 'clipboard-write']);
  await openLab(page);
  await selectMode(page, 'Reviewing');
  await page.evaluate(() => navigator.clipboard.writeText('2E-25\t=1-B2'));
  const editor = source(page);
  await editor.locator('td[data-x="1"][data-y="1"]').first().click();
  await page.keyboard.press('Control+V');
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveText('2e-25');
  await selectMode(page, 'Authoring');
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveText('1e-25');
  const proposal = editor.getByRole('article', { name: 'Suggestion: Paste cells', exact: true });
  await expect(proposal).toHaveCount(1);
  await expect(proposal).toContainText('2E-25');
  await expect(proposal).toContainText('=1-B2');
  await proposal.getByRole('button', { name: 'Accept suggestion', exact: true }).click();
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveText('2e-25');
});

test('SE-REVIEW-7: server calculations retain comments and accepted exact suggestions', async ({
  page,
}) => {
  await openLab(page);
  await page.getByRole('button', { name: 'Calculate on server', exact: true }).click();
  const editor = source(page);
  await expect(editor.getByRole('status', { name: 'Calculation status' })).toHaveText(
    'Calculated on server',
  );
  await selectMode(page, 'Reviewing');
  await editor.locator('td[data-x="1"][data-y="1"]').first().click();
  await editor.getByRole('textbox', { name: 'Cell value or formula', exact: true }).fill('2E-25');
  await editor
    .getByRole('textbox', { name: 'Cell value or formula', exact: true })
    .press('Control+Enter');
  const responsePromise = page.waitForResponse(
    (response) =>
      response.url().endsWith('/api/calculate') &&
      response.request().postDataJSON().workbook.worksheets[0].data[1][1] === '2E-25' &&
      response.request().postDataJSON().workbook.suggestions.at(-1).status === 'accepted',
  );
  await selectMode(page, 'Authoring');
  await editor
    .getByRole('article', { name: 'Suggestion: Edit B2', exact: true })
    .getByRole('button', { name: 'Accept suggestion', exact: true })
    .click();
  const response = await responsePromise;
  expect(response.ok()).toBe(true);
  expect(response.request().postDataJSON().workbook.comments).toEqual(
    expect.arrayContaining([
      expect.objectContaining({ range: 'B3' }),
      expect.objectContaining({
        state: 'resolved',
        decision: expect.objectContaining({ decision: 'accept' }),
      }),
    ]),
  );
  expect(response.request().postDataJSON().workbook.suggestions.at(-1).status).toBe('accepted');
  await expect(editor.locator('td[data-x="1"][data-y="3"]').first()).toHaveText('2e-25');
});

test('SE-CHROME-1: the history pane saves a named version and opens it read-only', async ({
  page,
}) => {
  await openLab(page);
  const editor = source(page);
  const formula = editor.getByRole('textbox', { name: 'Cell value or formula', exact: true });
  await formula.fill('5');
  await editor.getByRole('button', { name: 'History', exact: true }).click();
  await editor.getByRole('textbox', { name: 'Version name', exact: true }).fill('Review copy');
  await editor.getByRole('button', { name: 'Save version', exact: true }).click();
  const named = editor.getByText('Review copy', { exact: true });
  await expect(named).toBeVisible();
  await named.click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'View this version', exact: true }).click();
  await expect(editor.locator('.se-editor__historical')).toContainText('Review copy');
  await expect(formula).not.toBeEditable();
  await editor.getByRole('button', { name: 'View current version', exact: true }).click();
  await expect(formula).toBeEditable();
});
