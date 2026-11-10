import { expect, test } from '@playwright/test';

test('SE-LABS-3: index links to three independent editor configurations', async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('heading', { name: 'Spreadsheet editor labs' })).toBeVisible();
  for (const [name, chrome] of [
    ['Minimal', 'minimal'],
    ['Embedded in an app', 'embedded'],
    ['Full editor', 'full'],
  ]) {
    await page.getByRole('link', { name, exact: true }).click();
    const editor = page.locator('.se-editor').first();
    await expect(editor).toHaveAttribute('data-chrome', chrome);
    await expect(editor.locator('td[data-x="0"][data-y="1"]').first()).toBeVisible();
    await expect(editor.getByRole('menubar')).toHaveCount(chrome === 'full' ? 1 : 0);
    await expect(
      editor.getByRole('textbox', { name: 'Cell value or formula', exact: true }),
    ).toHaveCount(chrome === 'minimal' ? 0 : 1);
    await page.getByRole('link', { name: 'All labs', exact: true }).click();
  }
});

test('SE-LABS-3: stylesheet and theme choices follow navigation', async ({ page }) => {
  await page.goto('/?theme=dark');
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await page.getByRole('link', { name: 'Minimal', exact: true }).click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'dark');
  await page.getByRole('button', { name: /Stylesheet/ }).click();
  await page.getByRole('option', { name: 'Unstyled', exact: true }).click();
  expect(
    await page.locator('#lab-reference').evaluate((link) => (link as HTMLLinkElement).disabled),
  ).toBe(true);
  const cell = page.locator('td[data-x="1"][data-y="1"]').first();
  await cell.dblclick();
  await cell.getByRole('textbox').fill('8');
  await cell.getByRole('textbox').press('Enter');
  await expect(page.locator('td[data-x="3"][data-y="1"]').first()).toHaveText('4');
});

for (const width of [1440, 768, 390]) {
  test(`SE-CHROME-4: formula and toolbar layout stays within ${width}px`, async ({ page }) => {
    await page.setViewportSize({ width, height: 1000 });
    await page.goto('/full.html');
    const editor = page.locator('.se-editor').first();
    await expect(editor.locator('td[data-x="0"][data-y="0"]').first()).toBeVisible();
    const measurements = await editor.evaluate((root) => {
      const bar = root.querySelector('.se-editor__formula-bar')!.getBoundingClientRect();
      const toolbar = root.querySelector('.se-editor__toolbar')!.getBoundingClientRect();
      const input = root
        .querySelector<HTMLInputElement>('.se-editor__formula-input input')!
        .getBoundingClientRect();
      const apply = Array.from(root.querySelectorAll('button'))
        .find((button) => button.textContent === 'Apply')!
        .getBoundingClientRect();
      return {
        bar: bar.height,
        toolbar: toolbar.height,
        input: input.width,
        inputSpace: root.querySelector('.se-editor__formula-input')!.getBoundingClientRect().width,
        overlap: input.right > apply.left,
        overflow: document.documentElement.scrollWidth > innerWidth,
      };
    });
    expect(measurements.bar).toBeLessThan(75);
    expect(measurements.toolbar).toBeLessThan(48);
    expect(measurements.input).toBeGreaterThan(80);
    expect(measurements.input).toBeGreaterThan(measurements.inputSpace * 0.95);
    expect(measurements.overlap).toBe(false);
    expect(measurements.overflow).toBe(false);
    await editor
      .getByRole('textbox', { name: 'Cell value or formula', exact: true })
      .fill('=1-(1-1E-25)');
    await editor.getByRole('button', { name: 'Apply', exact: true }).click();
    await expect(editor.getByLabel('Calculated cell value')).toHaveText('1e-25');
  });
}

test('SE-CR-2.5: formula cancellation and reference picking preserve authored cells', async ({
  page,
}) => {
  await page.goto('/full.html');
  const editor = page.locator('.se-editor').first(),
    input = editor.getByRole('textbox', { name: 'Cell value or formula', exact: true });
  await input.fill('=');
  await editor.locator('td[data-x="1"][data-y="1"]').first().click();
  await expect(input).toHaveValue('=B2');
  await expect(editor.locator('.se-editor__reference-outline')).toHaveCount(1);
  await input.press('Escape');
  await expect(input).toHaveValue('Example');
  await expect(editor.locator('td[data-x="0"][data-y="0"]').first()).toHaveText('Example');
  await expect(editor.locator('.se-editor__reference-outline')).toHaveCount(0);
  await input.fill('=2+3');
  await input.press('Enter');
  await expect(editor.locator('td[data-x="0"][data-y="0"]').first()).toHaveText('5');
  await expect(
    editor.getByRole('textbox', { name: 'Selected cell or range', exact: true }),
  ).toHaveValue('A2');
});

test('SE-LABS-3: dark grid uses the active theme and compact controls show their values', async ({
  page,
}) => {
  await page.goto('/full.html?theme=dark');
  const editor = page.locator('.se-editor').first();
  const cell = editor.locator('td[data-x="1"][data-y="1"]').first();
  await expect(cell).toHaveCSS('background-color', 'rgb(21, 27, 35)');
  await expect(cell).toHaveCSS('color', 'rgb(240, 246, 252)');
  await expect(editor.locator('thead td[data-x="0"]').first()).toHaveCSS(
    'background-color',
    'rgb(13, 17, 23)',
  );
  await expect(editor.getByRole('button', { name: '12 Font size', exact: true })).toBeVisible();
  expect(
    await editor
      .locator('.se-editor__font-size .ui-select-field__value')
      .evaluate((element) => element.clientWidth),
  ).toBeGreaterThan(14);
});
