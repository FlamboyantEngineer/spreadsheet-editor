import { expect, test, type Locator, type Page } from '@playwright/test';

async function open(page: Page, lab = 'full', theme = 'light') {
  await page.goto(`/${lab}.html?theme=${theme}`);
  const editor = page.locator('.se-editor').first();
  await expect(editor.getByRole('status', { name: 'Calculation status' })).toHaveText(
    'Calculated on client',
  );
  return editor;
}

async function token(editor: Locator, name: string) {
  return editor.evaluate((root, name) => {
    const probe = document.createElement('span');
    probe.style.color = `var(${name})`;
    root.append(probe);
    const value = getComputedStyle(probe).color;
    probe.remove();
    return value;
  }, name);
}

test('SE-CONTEXT-1: cell context menus cancel the browser menu and retain usable commands', async ({
  page,
}) => {
  const editor = await open(page);
  await page.evaluate(() =>
    document.addEventListener(
      'contextmenu',
      (event) => {
        setTimeout(() => {
          document.documentElement.dataset.contextPrevented = String(event.defaultPrevented);
        });
      },
      { capture: true },
    ),
  );
  await editor.locator('td[data-x="1"][data-y="1"]').first().click({ button: 'right' });
  await expect(page.locator('html')).toHaveAttribute('data-context-prevented', 'true');
  const menu = page.getByRole('menu', { name: 'Spreadsheet context menu', exact: true });
  await expect(menu).toBeVisible();
  await menu.getByRole('menuitem', { name: 'Clear contents', exact: true }).click();
  await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveText('');
  const cell = editor.locator('td[data-x="1"][data-y="1"]').first();
  await cell.dblclick();
  await cell.getByRole('textbox').click({ button: 'right' });
  await expect(page.locator('html')).toHaveAttribute('data-context-prevented', 'false');
  await expect(menu).not.toBeVisible();
});

for (const theme of ['light', 'dark']) {
  test(`SE-CHROME-6: cells, headers and chrome use the shared background tokens in ${theme}`, async ({
    page,
  }) => {
    const editor = await open(page, 'full', theme);
    const canvas = await token(editor, '--ref-canvas');
    const surface = await token(editor, '--ref-surface');
    await expect(editor.locator('td[data-x="1"][data-y="1"]').first()).toHaveCSS(
      'background-color',
      canvas,
    );
    await expect(editor.locator('thead td[data-x="1"]').first()).toHaveCSS(
      'background-color',
      surface,
    );
    await expect(editor.locator('tbody td.jss_row[data-y="1"]').first()).toHaveCSS(
      'background-color',
      surface,
    );
    for (const selector of ['.se-editor__toolbar', '.se-editor__menubar'])
      await expect(editor.locator(selector)).toHaveCSS('background-color', canvas);
    const menu = editor.getByRole('menubar', { name: 'Spreadsheet menu' });
    await expect(menu).toHaveCSS('border-top-width', '0px');
    expect((await menu.boundingBox())!.width).toBeCloseTo(
      (await editor.boundingBox())!.width - 2,
      0,
    );
    await expect(editor.locator('.ui-formula-bar')).toHaveCSS('border-top-width', '0px');
    await expect(editor.locator('.ui-formula-bar')).toHaveCSS('border-left-width', '0px');
    await editor.getByRole('menuitem', { name: 'File', exact: true }).click();
    const popup = page.locator('.ui-menu-popover').filter({
      has: page.getByRole('menu', { name: 'File', exact: true }),
    });
    await expect(popup).toHaveCSS('background-color', await token(editor, '--ref-raised'));
  });
}

for (const lab of ['full', 'embedded']) {
  for (const width of [1440, 390]) {
    test(`SE-CHROME-6: ${lab} worksheet add and scrollbar stay inside their lanes at ${width}px`, async ({
      page,
    }) => {
      await page.setViewportSize({ width, height: 1000 });
      const editor = await open(page, lab);
      const tabs = editor.getByRole('tablist', { name: 'Worksheets', exact: true });
      const add = editor.getByRole('button', { name: 'Add worksheet', exact: true });
      await add.scrollIntoViewIfNeeded();
      const lastBox = (await editor.getByRole('tab').last().boundingBox())!;
      const addBox = (await add.boundingBox())!;
      expect(addBox.x).toBeGreaterThanOrEqual(lastBox.x + lastBox.width);
      expect(addBox.x - lastBox.x - lastBox.width).toBeLessThan(12);
      await expect(tabs).toHaveAttribute('aria-label', 'Worksheets');
      await expect(
        editor.getByRole('tab', { name: lab === 'full' ? 'Inputs' : 'Mass budget', exact: true }),
      ).toBeVisible();
      const measure = () =>
        editor.evaluate((root) => {
          const grid = root.querySelector<HTMLElement>('.jss_content')!;
          const sidebar = root.querySelector<HTMLElement>('.se-editor__sidebar')!;
          const box = grid.getBoundingClientRect(),
            rail = sidebar.getBoundingClientRect();
          const target = document.elementFromPoint(
            box.right - 2,
            box.top + Math.min(100, box.height / 2),
          );
          return {
            gutter: grid.offsetWidth - grid.clientWidth,
            edge: box.right,
            rail: rail.left,
            covered: !!target && sidebar.contains(target),
            overflow: document.documentElement.scrollWidth > innerWidth,
          };
        });
      const initial = await measure();
      expect(initial.gutter).toBeGreaterThanOrEqual(4);
      expect(initial.edge).toBeLessThanOrEqual(initial.rail);
      expect(initial.covered).toBe(false);
      expect(initial.overflow).toBe(false);
      const viewport = editor.locator('.jss_content').first();
      await viewport.evaluate((element) => {
        element.scrollTop = 400;
      });
      await expect
        .poll(() => viewport.evaluate((element) => element.scrollTop))
        .toBeGreaterThan(100);
      const zoom = editor.getByRole('group', { name: 'Spreadsheet zoom', exact: true });
      await zoom.getByRole('button', { name: 'Zoom in', exact: true }).click();
      expect((await measure()).covered).toBe(false);
      if (width === 1440) {
        await editor.getByRole('button', { name: 'Properties', exact: true }).click();
        await expect(editor.locator('.ui-dock-panel')).toHaveAttribute('data-expanded');
        const expanded = await measure();
        expect(expanded.covered).toBe(false);
        expect(expanded.edge).toBeLessThanOrEqual(expanded.rail);
        const panel = editor.locator('.se-editor__sidebar');
        const before = (await panel.boundingBox())!.width;
        const handle = editor.getByRole('separator', {
          name: 'Resize Spreadsheet panels',
          exact: true,
        });
        const rect = (await handle.boundingBox())!;
        await page.mouse.move(rect.x + rect.width / 2, rect.y + rect.height / 2);
        await page.mouse.down();
        await page.mouse.move(rect.x + rect.width / 2 - 30, rect.y + rect.height / 2);
        await page.mouse.up();
        await expect
          .poll(async () => (await panel.boundingBox())!.width)
          .toBeGreaterThan(before + 20);
        expect((await measure()).covered).toBe(false);
      }
    });
  }
}

test('SE-CHROME-6: the full toolbar has one row and icon actions, with readable field controls', async ({
  page,
}) => {
  const editor = await open(page);
  const toolbar = editor.getByRole('toolbar', { name: 'Spreadsheet formatting', exact: true });
  await expect(editor.locator('.se-editor__chrome').getByRole('toolbar')).toHaveCount(1);
  await expect(editor.locator('.se-editor__actions')).toHaveCount(0);
  for (const name of [
    'Wrap text',
    'Merge cells',
    'Unmerge cells',
    'Align cells',
    'Borders',
    'Find',
    'Add row',
    'Add column',
    'Comments (1)',
    'Suggestions (1)',
  ]) {
    const button = toolbar.getByRole('button', { name, exact: true });
    await expect(button.locator('svg')).toHaveCount(1);
    expect((await button.innerText()).trim()).toBe('');
  }
  await expect(toolbar.getByRole('button', { name: /Font family$/ })).toBeVisible();
  await expect(toolbar.getByRole('button', { name: /Font size$/ })).toBeVisible();
});

test('SE-CHROME-6: the embedded menu uses the public three-line glyph with an intact button border', async ({
  page,
}) => {
  const editor = await open(page, 'embedded');
  const trigger = editor.getByRole('button', { name: 'Spreadsheet actions', exact: true });
  await expect(trigger).toHaveAttribute('data-variant', 'secondary');
  const borders = await trigger.evaluate((element) => {
    const style = getComputedStyle(element);
    return [
      style.borderTopWidth,
      style.borderRightWidth,
      style.borderBottomWidth,
      style.borderLeftWidth,
    ];
  });
  expect(borders).toEqual(['1px', '1px', '1px', '1px']);
  await expect(trigger.locator('svg')).toHaveCount(1);
  await trigger.click();
  await expect(page.getByRole('menuitem', { name: 'File', exact: true })).toBeVisible();
});
