import { expect, test } from '@playwright/test';
import {
  clickCell,
  enter,
  expectCell,
  isCellPainted,
  openApp,
  openStoredSheet,
} from './grid';

// Also the smoke test that the deployment runs against the live address.
test('a formula typed in the grid shows its result @smoke', async ({
  page,
}) => {
  test.setTimeout(150_000); // A new deployment can take a while to be served.
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '2'); // Enter moves down to A2.
  await enter(page, '=A1+1');

  await expectCell(page, 1, 0, 'A2 3');
  await expectCell(page, 0, 0, 'A1 2');

  // The result is drawn on the canvas, in A2 and not beside it. The
  // selection is moved away first so that its border is not what is seen.
  await clickCell(page, 5, 3);
  expect(await isCellPainted(page, 1, 0)).toBe(true);
  expect(await isCellPainted(page, 1, 1)).toBe(false);
});

test('Escape abandons an entry', async ({ page }) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await page.keyboard.type('7');
  await page.keyboard.press('Escape');
  await expectCell(page, 0, 0, 'A1');
});

test('formulas combine operators, ranges and functions, and show errors', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '2');
  await enter(page, '3');
  await enter(page, '=SUM(A1:A2)*2^2');
  await enter(page, '=IF(A3>10,"big","small")&"!"');
  await enter(page, '=A1/0');
  await enter(page, '=A5+1');

  await expectCell(page, 2, 0, 'A3 20');
  await expectCell(page, 3, 0, 'A4 big!');
  await expectCell(page, 4, 0, 'A5 #DIV/0!');
  await expectCell(page, 5, 0, 'A6 #DIV/0!');

  // A change reaches the formulas that depend on it, directly or not.
  await clickCell(page, 0, 0);
  await enter(page, '0.5');
  await expectCell(page, 2, 0, 'A3 14');
  await expectCell(page, 3, 0, 'A4 big!');
});

test('functions of each family give their results, and errors for wrong arguments', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '10');
  await enter(page, '20');
  await enter(page, '=SUMIF(A1:A2,">15")+ROUND(PI(),2)');
  await enter(page, '=IFERROR(STDEV.S(A1:A2)/0,"none")');
  await enter(page, '=SQRT("x")');
  await enter(page, '=ROUND(1)');

  await expectCell(page, 2, 0, 'A3 23.14');
  await expectCell(page, 3, 0, 'A4 none');
  await expectCell(page, 4, 0, 'A5 #VALUE!');
  // Too few arguments: not a formula, so kept as text, as Excel refuses it.
  await expectCell(page, 5, 0, 'A6 =ROUND(1)');
});

test('a circular reference is reported, and the app keeps working', async ({
  page,
}) => {
  await openApp(page);
  const statusBar = page.getByRole('contentinfo');
  await clickCell(page, 0, 0);
  await enter(page, '=B1+1');
  await clickCell(page, 0, 1);
  await enter(page, '=A1+1');
  await expect(statusBar).toContainText('Circular References: B1');
  await expectCell(page, 0, 0, 'A1 0');

  await clickCell(page, 0, 1);
  await enter(page, '5');
  await expect(statusBar).not.toContainText('Circular References');
  await expectCell(page, 0, 0, 'A1 6');
});

test('a change that 100,000 formulas depend on shows within a second, and input goes on', async ({
  page,
}) => {
  test.setTimeout(60_000);
  await openApp(page);
  const rows = 100_000;
  await openStoredSheet(page, {
    columns: [[1]],
    formulas: Array.from({ length: rows }, (_, row) => [
      row,
      1,
      row === 0 ? '=A1+1' : `=B${row}+1`,
    ]),
  });
  await expectCell(page, 0, 1, `B1 2`);

  await clickCell(page, 0, 0);
  await page.keyboard.type('2');
  const start = Date.now();
  await page.keyboard.press('Enter');
  // Typed at once, while the change is still taking effect.
  await enter(page, '7');
  await expectCell(page, 0, 1, 'B1 3');
  expect(Date.now() - start).toBeLessThan(1000);
  await expectCell(page, 1, 0, 'A2 7');

  // The end of the chain, which depends on every formula before it.
  await clickCell(page, 0, 1);
  await page.keyboard.press('ControlOrMeta+ArrowDown');
  await expect(page.getByRole('status')).toHaveText(`B${rows} ${rows + 2}`);
});

test('text, date, lookup and information functions give their results', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, 'pear');
  await enter(page, 'apple');
  await clickCell(page, 0, 1);
  await enter(page, '3');
  await enter(page, '5');
  await clickCell(page, 0, 2);
  await enter(page, '=VLOOKUP("apple",A1:B2,2,FALSE)&"/"&PROPER(A1)');
  await enter(page, '=TEXT(DATE(2024,2,29)+1,"yyyy-mm-dd")');
  await enter(page, '=XLOOKUP("kiwi",A1:A2,B1:B2)');
  await enter(page, '=ISNA(C3)');

  await expectCell(page, 0, 2, 'C1 5/Pear');
  await expectCell(page, 1, 2, 'C2 2024-03-01');
  await expectCell(page, 2, 2, 'C3 #N/A');
  await expectCell(page, 3, 2, 'C4 TRUE');
});
