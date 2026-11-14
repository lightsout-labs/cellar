import { test } from '@playwright/test';
import { clickCell, enter, expectCell, openApp } from './grid';

test('cells are restored from browser storage after a reload', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '2');
  await enter(page, '=A1+1');
  await expectCell(page, 1, 0, 'A2 3');

  await page.reload();
  await expectCell(page, 1, 0, 'A2 3');
  await expectCell(page, 0, 0, 'A1 2');

  // The formula itself was stored, not only its result.
  await clickCell(page, 0, 0);
  await enter(page, '10');
  await expectCell(page, 1, 0, 'A2 11');
});

test('every edit is present after closing and reopening the browser', async ({
  browserName,
  playwright,
  baseURL,
}, testInfo) => {
  const userData = testInfo.outputPath('profile');
  const launch = () =>
    playwright[browserName].launchPersistentContext(userData, { baseURL });

  let browser = await launch();
  let page = browser.pages()[0] ?? (await browser.newPage());
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, '1/2/2026');
  await enter(page, 'last');
  // Closed at once, without waiting for anything to be saved.
  await browser.close();

  browser = await launch();
  page = browser.pages()[0] ?? (await browser.newPage());
  await openApp(page);
  await expectCell(page, 0, 0, 'A1 1/2/2026');
  await expectCell(page, 1, 0, 'A2 last');
  await browser.close();
});
