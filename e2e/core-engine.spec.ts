import { expect, test } from '@playwright/test';

test('SE-UI-1: final blank rows and columns synchronize with the live vendor without replacing its root', async ({
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

test('SE-UI-1: selection changes write no frozen styles on an unchanged 26 by 100 grid', async ({
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
      const descriptor = Object.getOwnPropertyDescriptor(CSSStyleDeclaration.prototype, property)!;
      Object.defineProperty(CSSStyleDeclaration.prototype, property, {
        ...descriptor,
        set(value: string) {
          if (styles.has(this)) counter.styles++;
          descriptor.set!.call(this, value);
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
