import { expect, test, type Page } from '@playwright/test';
import { clickCell, enter, expectCell, openApp } from './grid';

const status = (page: Page) => page.getByRole('status');
const summary = (page: Page) => page.getByLabel('Summary of the selection');

test.use({ locale: 'en-US' });

test('undo and redo with the keyboard and the ribbon, selecting what changed', async ({
  page,
}) => {
  await openApp(page);
  const undo = page.getByRole('button', { name: 'Undo' });
  const redo = page.getByRole('button', { name: 'Redo' });
  await expect(undo).toBeDisabled();
  await expect(redo).toBeDisabled();

  await clickCell(page, 0, 0);
  await enter(page, '1');
  await enter(page, '2');
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Delete');
  await clickCell(page, 5, 5);

  // Undo brings back the cleared cells and selects them.
  await page.keyboard.press('ControlOrMeta+z');
  await expect(status(page)).toHaveText('A1 1');
  await expect(summary(page)).toHaveText('Average: 1.5Count: 2Sum: 3');
  await expect(redo).toBeEnabled();

  // The ribbon undoes the entry in A2, and the keyboard still works after.
  await undo.click();
  await expect(status(page)).toHaveText('A2');
  await page.keyboard.press('ControlOrMeta+y');
  await expect(status(page)).toHaveText('A2 2');
  await redo.click();
  await expect(status(page)).toHaveText('A1');
  await expect(redo).toBeDisabled();

  // A new change after an undo leaves nothing to redo.
  await page.keyboard.press('ControlOrMeta+z');
  await expect(redo).toBeEnabled();
  await clickCell(page, 0, 2);
  await enter(page, '3');
  await expect(redo).toBeDisabled();

  // What undo did is stored: it is still there after a reload.
  await page.keyboard.press('ControlOrMeta+z');
  await expect(status(page)).toHaveText('C1');
  // A reload discards a write to storage still under way, as S01 found;
  // this waits for the autosave, half a second after the change.
  await page.waitForTimeout(1000);
  await page.reload();
  await expect(page.locator('canvas')).toBeVisible();
  await expectCell(page, 0, 0, 'A1 1');
  await expectCell(page, 1, 0, 'A2 2');
  await expectCell(page, 0, 2, 'C1');
  await expect(undo).toBeDisabled();
});

test('during an entry, the shortcut leaves the workbook alone', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1');
  await clickCell(page, 0, 0);
  await page.keyboard.press('F2');
  await page.keyboard.press('ControlOrMeta+z');
  await page.keyboard.press('Escape');
  await expect(status(page)).toHaveText('A1 1');
});
