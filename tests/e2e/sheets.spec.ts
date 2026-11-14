import { expect, test, type Locator, type Page } from '@playwright/test';
import { clickCell, enter, expectCell, openApp } from './grid';

test.use({ locale: 'en-US' });

const tabs = (page: Page) =>
  page.getByRole('tablist', { name: 'Sheets' }).getByRole('tab');
const tab = (page: Page, name: string) =>
  page.getByRole('tab', { name, exact: true });

/** Opens the context menu of a tab and runs one of its actions. */
async function menu(page: Page, name: string, action: string) {
  await tab(page, name).click({ button: 'right' });
  await page.getByRole('menuitem', { name: action, exact: true }).click();
}

/** Drags a tab by the mouse and drops it on the far side of another. */
async function drag(page: Page, from: Locator, to: Locator) {
  const start = (await from.boundingBox())!;
  const end = (await to.boundingBox())!;
  await page.mouse.move(start.x + start.width / 2, start.y + start.height / 2);
  await page.mouse.down();
  await page.mouse.move(end.x + end.width / 2, end.y + end.height / 2, {
    steps: 5,
  });
  const far = end.x > start.x ? end.x + end.width - 2 : end.x + 2;
  await page.mouse.move(far, end.y + end.height / 2, { steps: 2 });
  await page.mouse.up();
}

test('add, rename, duplicate, reorder, color and delete sheets, which a reload keeps', async ({
  page,
}) => {
  await openApp(page);
  await expect(tabs(page)).toHaveText(['Sheet1']);
  await clickCell(page, 0, 0);
  await enter(page, 'one');

  // The new sheet is shown, empty.
  await page.getByRole('button', { name: 'New sheet' }).click();
  await expect(tabs(page)).toHaveText(['Sheet1', 'Sheet2']);
  await expect(tab(page, 'Sheet2')).toHaveAttribute('aria-selected', 'true');
  await expectCell(page, 0, 0, 'A1');

  // Double-clicking a tab renames it.
  await tab(page, 'Sheet2').dblclick();
  const name = page.getByRole('textbox', { name: 'Sheet name' });
  await name.fill('Sales');
  await name.press('Enter');
  await expect(tabs(page)).toHaveText(['Sheet1', 'Sales']);

  // The context menu duplicates a sheet with its contents.
  await menu(page, 'Sheet1', 'Duplicate');
  await expect(tabs(page)).toHaveText(['Sheet1', 'Sheet1 (2)', 'Sales']);
  await expectCell(page, 0, 0, 'A1 one');

  await menu(page, 'Sales', 'Rename');
  await name.fill('Costs');
  await name.press('Enter');
  await menu(page, 'Sheet1', 'Insert');
  await expect(tabs(page)).toHaveText([
    'Sheet1',
    'Sheet4',
    'Sheet1 (2)',
    'Costs',
  ]);
  await menu(page, 'Sheet4', 'Delete');
  await expect(tabs(page)).toHaveText(['Sheet1', 'Sheet1 (2)', 'Costs']);

  // Dragging Sheet1 past Costs moves it to the end.
  await drag(page, tab(page, 'Sheet1'), tab(page, 'Costs'));
  await expect(tabs(page)).toHaveText(['Sheet1 (2)', 'Costs', 'Sheet1']);

  await tab(page, 'Costs').click({ button: 'right' });
  await page.getByRole('menuitem', { name: 'Tab color' }).click();
  await page.getByRole('menuitem', { name: 'Red' }).click();
  await expect(tab(page, 'Costs')).toHaveCSS(
    'box-shadow',
    /rgb\(229, 72, 77\)/,
  );

  // Each sheet keeps its own cells.
  await tab(page, 'Sheet1').click();
  await expectCell(page, 0, 0, 'A1 one');
  await tab(page, 'Costs').click();
  await expectCell(page, 0, 0, 'A1');

  // A reload discards a write still under way (S01): wait for the autosave.
  await page.waitForTimeout(1000);
  await page.reload();
  await expect(page.locator('canvas')).toBeVisible();
  await expect(tabs(page)).toHaveText(['Sheet1 (2)', 'Costs', 'Sheet1']);
  await expect(tab(page, 'Costs')).toHaveCSS(
    'box-shadow',
    /rgb\(229, 72, 77\)/,
  );
  await tab(page, 'Sheet1').click();
  await expectCell(page, 0, 0, 'A1 one');
});

test('each change to the sheets can be undone and redone', async ({ page }) => {
  await openApp(page);
  await page.getByRole('button', { name: 'New sheet' }).click();
  await menu(page, 'Sheet2', 'Rename');
  const name = page.getByRole('textbox', { name: 'Sheet name' });
  await name.fill('Plan');
  await name.press('Enter');
  await drag(page, tab(page, 'Plan'), tab(page, 'Sheet1'));
  await expect(tabs(page)).toHaveText(['Plan', 'Sheet1']);
  await menu(page, 'Sheet1', 'Delete');
  await expect(tabs(page)).toHaveText(['Plan']);

  const undo = page.getByRole('button', { name: 'Undo' });
  await undo.click();
  await expect(tabs(page)).toHaveText(['Plan', 'Sheet1']);
  await expect(tab(page, 'Sheet1')).toHaveAttribute('aria-selected', 'true');
  await page.keyboard.press('ControlOrMeta+z');
  await expect(tabs(page)).toHaveText(['Sheet1', 'Plan']);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(tabs(page)).toHaveText(['Sheet1', 'Sheet2']);
  await page.keyboard.press('ControlOrMeta+z');
  await expect(tabs(page)).toHaveText(['Sheet1']);
  await page.keyboard.press('ControlOrMeta+y');
  await page.keyboard.press('ControlOrMeta+y');
  await page.keyboard.press('ControlOrMeta+y');
  await page.keyboard.press('ControlOrMeta+y');
  await expect(tabs(page)).toHaveText(['Plan']);
});

test('the last sheet cannot be deleted, and a name Excel refuses is refused', async ({
  page,
}) => {
  await openApp(page);
  await tab(page, 'Sheet1').click({ button: 'right' });
  await expect(
    page.getByRole('menuitem', { name: 'Delete', exact: true }),
  ).toBeDisabled();
  await page.keyboard.press('Escape');

  await page.getByRole('button', { name: 'New sheet' }).click();
  await tab(page, 'Sheet2').dblclick();
  const name = page.getByRole('textbox', { name: 'Sheet name' });
  await name.fill('sheet1');
  await name.press('Enter');
  await expect(page.getByRole('alert')).toHaveText(
    'Another sheet already has this name.',
  );
  await expect(name).toBeVisible();
  // Escape keeps the name it had.
  await name.press('Escape');
  await expect(tabs(page)).toHaveText(['Sheet1', 'Sheet2']);
});
