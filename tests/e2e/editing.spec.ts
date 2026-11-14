import { expect, test, type Page } from '@playwright/test';
import {
  clickCell,
  doubleClickCell,
  enter,
  expectCell,
  isCellPainted,
  openApp,
} from './grid';

const status = (page: Page) => page.getByRole('status');

test('typing, double-clicking, F2 and the formula bar each edit a cell', async ({
  page,
}) => {
  await openApp(page);

  // Typing replaces the contents; Tab commits and moves right.
  await clickCell(page, 0, 0);
  await page.keyboard.type('abc');
  await page.keyboard.press('Tab');
  await expect(status(page)).toHaveText('B1');
  await expectCell(page, 0, 0, 'A1 abc');

  // F2 edits the contents, the caret at the end; Enter commits and moves down.
  await page.keyboard.press('F2');
  await expect(page.getByRole('contentinfo')).toContainText('Edit');
  await page.keyboard.type('d');
  await page.keyboard.press('Enter');
  await expect(status(page)).toHaveText('A2');
  await expectCell(page, 0, 0, 'A1 abcd');

  // Double-clicking edits too; Escape abandons the change.
  await doubleClickCell(page, 0, 0);
  await expect(page.getByRole('textbox', { name: 'A1' })).toHaveValue('abcd');
  await page.keyboard.type('e');
  await page.keyboard.press('Escape');
  await expect(status(page)).toHaveText('A1 abcd');

  // The formula bar edits the active cell.
  const formulaBar = page.getByRole('textbox', { name: 'Formula bar' });
  await expect(formulaBar).toHaveValue('abcd');
  await formulaBar.click();
  await formulaBar.fill('=1+2');
  await expect(page.getByRole('textbox', { name: 'A1' })).toHaveValue('=1+2');
  await page.keyboard.press('Enter');
  await expect(status(page)).toHaveText('A2');
  await expectCell(page, 0, 0, 'A1 3');
  await expect(formulaBar).toHaveValue('=1+2');

  // In an entry started by typing, an arrow commits and moves.
  await clickCell(page, 2, 2);
  await page.keyboard.type('5');
  await page.keyboard.press('ArrowRight');
  await expect(status(page)).toHaveText('D3');
  await expectCell(page, 2, 2, 'C3 5');
});

test('typed numbers, dates, times, percentages and booleans are recognized', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  for (const text of [
    '1/2/2026',
    '13:45',
    '12%',
    'true',
    '1,234.5',
    'hello',
    '=A3+1',
  ]) {
    await enter(page, text);
  }
  await expectCell(page, 0, 0, 'A1 1/2/2026');
  await expectCell(page, 1, 0, 'A2 13:45');
  await expectCell(page, 2, 0, 'A3 12%');
  await expectCell(page, 3, 0, 'A4 TRUE');
  await expectCell(page, 4, 0, 'A5 1,234.50');
  await expectCell(page, 5, 0, 'A6 hello');
  // The percentage is the number 0.12.
  await expectCell(page, 6, 0, 'A7 1.12');
  // A date is edited as a date.
  await clickCell(page, 0, 0);
  await expect(page.getByRole('textbox', { name: 'Formula bar' })).toHaveValue(
    '1/2/2026',
  );
});

test('numbers align right, text left, and long text runs over empty cells', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '123');
  await enter(page, 'abc');
  await enter(page, 'a long text that runs over');
  await clickCell(page, 3, 0);
  await enter(page, 'a long text that is cut');
  await clickCell(page, 3, 1);
  await enter(page, '7');
  await clickCell(page, 10, 5); // The selection is not what is seen.

  expect(await isCellPainted(page, 0, 0, 'left')).toBe(false);
  expect(await isCellPainted(page, 0, 0, 'right')).toBe(true);
  expect(await isCellPainted(page, 1, 0, 'left')).toBe(true);
  expect(await isCellPainted(page, 1, 0, 'right')).toBe(false);
  expect(await isCellPainted(page, 2, 1)).toBe(true);
  // B4 holds a number, at its right, and the text of A4 stops before it.
  expect(await isCellPainted(page, 3, 1, 'left')).toBe(false);
});

test('Delete clears the selected cells', async ({ page }) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1');
  await enter(page, '2');
  await clickCell(page, 0, 1);
  await enter(page, 'x');
  await clickCell(page, 0, 0);
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Delete');
  await expectCell(page, 0, 0, 'A1');
  await expectCell(page, 1, 0, 'A2');
  await expectCell(page, 0, 1, 'B1');
});
