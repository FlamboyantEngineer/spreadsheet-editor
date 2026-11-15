import { expect, test } from '@playwright/test';
import { requireDefined } from '../src/model/require-defined.js';

test('SE-UI-3.1: delayed initial grid adoption does not interrupt a held toolbar press', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1440, height: 1400 });
  let releaseChunk!: () => void;
  const released = new Promise<void>((resolve) => {
    releaseChunk = resolve;
  });
  let markRequested!: () => void;
  const requested = new Promise<void>((resolve) => {
    markRequested = resolve;
  });
  await page.route(/\/workbook-grid-[^/]+\.js$/, async (route) => {
    markRequested();
    await released;
    await route.continue();
  });
  await page.goto('/embedded.html', { waitUntil: 'domcontentloaded' });
  await requested;
  const editor = page.getByRole('region', { name: 'Embedded spreadsheet', exact: true });
  const add = editor.getByRole('button', { name: 'Add worksheet', exact: true });
  await add.scrollIntoViewIfNeeded();
  const box = requireDefined(await add.boundingBox());
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  releaseChunk();
  await expect(editor.locator('[data-jspreadsheet-host]')).toBeAttached();
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
  await expect(add).toBeFocused();
  const mountedBox = requireDefined(await add.boundingBox());
  expect(mountedBox).toEqual(box);
  await page.mouse.up();
  await expect(editor.getByRole('tab')).toHaveCount(2);
  await expect(editor.locator('td[data-x="0"][data-y="0"]').first()).toHaveText('');
});

test('SE-UI-3.1: a refused native batch restores accepted cells and can be retried', async ({
  page,
}) => {
  await page.goto('/host.html');
  const editor = page.getByRole('region', { name: 'Host spreadsheet', exact: true });
  const cells = [0, 1, 2].map((x) => editor.locator(`td[data-x="${x}"][data-y="0"]`).first());
  await expect(cells[1]).toHaveText('6');
  const original = await Promise.all(cells.map((cell) => cell.textContent()));
  await page.getByRole('button', { name: 'Reject workbook changes', exact: true }).click();
  await cells[0].click();
  await page.keyboard.down('Shift');
  await cells[2].click();
  await page.keyboard.up('Shift');
  await page.keyboard.press('Delete');
  for (let x = 0; x < cells.length; x++)
    await expect(cells[x]).toHaveText(requireDefined(original[x]));
  await expect(editor.getByRole('alert')).toContainText('Host workbook change was rejected.');
  await expect(page.getByLabel('Host transaction count')).toHaveText('0');
  await page.getByRole('button', { name: 'Allow workbook changes', exact: true }).click();
  await cells[0].click();
  await page.keyboard.down('Shift');
  await cells[2].click();
  await page.keyboard.up('Shift');
  await page.keyboard.press('Delete');
  for (const cell of cells) await expect(cell).toHaveText('');
  await expect(page.getByLabel('Host transaction count')).toHaveText('1');
});

for (const key of ['Delete', 'Backspace'])
  test(`SE-UI-1.1: ${key} clears a native multi-cell selection in one undo step`, async ({
    page,
  }) => {
    await page.goto('/full.html');
    const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
    const cells = [0, 1, 2].map((x) => editor.locator(`td[data-x="${x}"][data-y="0"]`).first());
    const original = await Promise.all(cells.map((cell) => cell.textContent()));
    await cells[0].click();
    await page.keyboard.down('Shift');
    await cells[2].click();
    await page.keyboard.up('Shift');
    await page.keyboard.press(key);
    for (const cell of cells) await expect(cell).toHaveText('');
    await expect(editor.locator('.se-editor__notice')).toHaveCount(0);
    await editor.getByRole('button', { name: 'Undo', exact: true }).click();
    for (let x = 0; x < cells.length; x++)
      await expect(cells[x]).toHaveText(requireDefined(original[x]));
  });

test('SE-UI-3.1: native bulk clear creates one review proposal without changing accepted cells', async ({
  page,
}) => {
  await page.goto('/full.html');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  const cells = [0, 1, 2].map((x) => editor.locator(`td[data-x="${x}"][data-y="0"]`).first());
  const original = await Promise.all(cells.map((cell) => cell.textContent()));
  await page.getByRole('button', { name: 'Reviewer permissions', exact: true }).click();
  await cells[0].click();
  await page.keyboard.down('Shift');
  await cells[2].click();
  await page.keyboard.up('Shift');
  await page.keyboard.press('Delete');
  for (const cell of cells) await expect(cell).toHaveText('');
  await expect(editor.locator('.se-editor__notice')).toHaveCount(0);
  await page.getByRole('button', { name: 'Reviewer permissions', exact: true }).click();
  await editor.getByRole('button', { name: /Editing mode$/ }).click();
  await page.getByRole('option', { name: 'Authoring', exact: true }).click();
  for (let x = 0; x < cells.length; x++)
    await expect(cells[x]).toHaveText(requireDefined(original[x]));
});

test('SE-UI-1.1: final blank rows and columns synchronize with the live vendor without replacing its root', async ({
  page,
}) => {
  await page.goto('/full.html');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  const engine = editor.locator('[data-jspreadsheet-host]').first();
  await expect(editor.locator('td[data-x="25"][data-y="99"]').first()).toBeAttached();
  await engine.evaluate((element) => element.setAttribute('data-core-stable-root', 'true'));
  await editor.locator('td[data-x="25"][data-y="99"]').first().click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Insert rows below', exact: true }).click();
  await expect(editor.locator('td[data-x="25"][data-y="100"]').first()).toBeAttached();
  await editor.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('td[data-x="25"][data-y="100"]')).toHaveCount(0);
  await editor.locator('td[data-x="25"][data-y="0"]').first().click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Insert columns after', exact: true }).click();
  await expect(editor.locator('td[data-x="26"][data-y="0"]').first()).toBeAttached();
  await editor.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(editor.locator('td[data-x="26"][data-y="0"]')).toHaveCount(0);
  await expect(engine).toHaveAttribute('data-core-stable-root', 'true');
});

test('SE-UI-1.1: selection changes write no frozen styles on an unchanged 26 by 100 grid', async ({
  page,
}) => {
  await page.goto('/full.html');
  const editor = page.getByRole('region', { name: 'Spreadsheet editor', exact: true });
  await editor.getByRole('button', { name: 'Add worksheet', exact: true }).click();
  await editor.locator('td[data-x="0"][data-y="0"]').first().click();
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
  await editor.evaluate((element) => {
    const cells = Array.from(element.querySelectorAll<HTMLElement>('.se-editor__engine td'));
    const styles = new Set(cells.map((cell) => cell.style));
    const classes = new Set(cells.map((cell) => cell.classList));
    const counter = { styles: 0, classes: 0 };
    (window as unknown as { coreSelectionWrites: typeof counter }).coreSelectionWrites = counter;
    for (const property of ['left', 'position', 'zIndex']) {
      const descriptor =
        Object.getOwnPropertyDescriptor(CSSStyleDeclaration.prototype, property) ?? {};
      Object.defineProperty(CSSStyleDeclaration.prototype, property, {
        ...descriptor,
        set(value: string) {
          if (styles.has(this)) counter.styles++;
          descriptor.set?.call(this, value);
        },
      });
    }
    for (const method of ['add', 'remove', 'toggle'] as const) {
      const original = DOMTokenList.prototype[method];
      Object.defineProperty(DOMTokenList.prototype, method, {
        configurable: true,
        value(this: DOMTokenList, ...tokens: unknown[]) {
          if (classes.has(this) && tokens[0] === 'jss_freezed') counter.classes++;
          return Reflect.apply(original, this, tokens);
        },
      });
    }
  });
  const start = Date.now();
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Shift+ArrowDown');
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
  const writes = await page.evaluate(
    () =>
      (window as unknown as { coreSelectionWrites: { styles: number; classes: number } })
        .coreSelectionWrites,
  );
  expect(writes).toEqual({ styles: 0, classes: 0 });
  test.info().annotations.push({
    type: 'diagnostic',
    description: `Selection and two paint frames: ${Date.now() - start} ms`,
  });
});
