import { readFile } from 'node:fs/promises';
import { expect, test, type Page } from '@playwright/test';
import { UNSUPPORTED_PARTS, xlsxFile } from '../xlsx';
import { clickCell, enter, expectCell, openApp } from './grid';

const XLSX_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const sheetTabs = (page: Page) =>
  page.getByRole('tablist', { name: 'Sheets' }).getByRole('tab');

/** Opens a file with the Open button of the File tab. */
async function openFile(
  page: Page,
  name: string,
  bytes: Uint8Array,
): Promise<void> {
  await page.getByRole('tab', { name: 'File' }).click();
  const chooser = page.waitForEvent('filechooser');
  await page.getByRole('button', { name: 'Open' }).click();
  await (
    await chooser
  ).setFiles({ name, mimeType: XLSX_TYPE, buffer: Buffer.from(bytes) });
}

const budget = () =>
  xlsxFile({
    strings: ['<t>Rent</t>'],
    sheets: [
      {
        name: 'Income',
        rows:
          '<row r="1"><c r="A1"><v>5</v></c></row>' +
          '<row r="2"><c r="A2"><f>A1*2</f><v>10</v></c></row>',
      },
      { name: 'Costs', rows: '<row r="1"><c r="A1" t="s"><v>0</v></c></row>' },
    ],
  });

test('an .xlsx file opens, and saving it gives a file that opens the same', async ({
  page,
}) => {
  await openApp(page);
  await openFile(page, 'Budget.xlsx', await budget());

  await expect(sheetTabs(page)).toHaveText(['Income', 'Costs']);
  await expectCell(page, 0, 0, 'A1 5');
  await expectCell(page, 1, 0, 'A2 10');
  await expect(page.getByLabel('Formula bar')).toHaveValue('=A1*2');
  await expect(page.getByRole('alert')).toHaveCount(0);

  // The workbook opened is the one kept in browser storage.
  await page.reload();
  await expect(sheetTabs(page)).toHaveText(['Income', 'Costs']);
  await clickCell(page, 0, 0);
  await enter(page, '7');
  await expectCell(page, 1, 0, 'A2 14');

  await page.getByRole('tab', { name: 'File' }).click();
  const downloading = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Save' }).click();
  const download = await downloading;
  expect(download.suggestedFilename()).toBe('Budget.xlsx');
  const saved = await readFile(await download.path());

  // A change after the save is not in the file saved.
  await clickCell(page, 0, 1);
  await enter(page, 'later');
  await openFile(page, 'Budget.xlsx', saved);
  await expect(sheetTabs(page)).toHaveText(['Income', 'Costs']);
  await expectCell(page, 0, 1, 'B1');
  await expectCell(page, 0, 0, 'A1 7');
  await expectCell(page, 1, 0, 'A2 14');
  await expect(page.getByLabel('Formula bar')).toHaveValue('=A1*2');
  await sheetTabs(page).filter({ hasText: 'Costs' }).click();
  await expectCell(page, 0, 0, 'A1 Rent');
});

test('a file with features Gridworks does not support opens and says what was dropped', async ({
  page,
}) => {
  await openApp(page);
  await openFile(
    page,
    'Report.xlsx',
    await xlsxFile({
      sheets: [
        {
          name: 'Report',
          rows: '<row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>_xlfn.XLOOKUP(A1,C1:C2,D1:D2)</f><v>9</v></c></row>',
        },
      ],
      parts: UNSUPPORTED_PARTS,
    }),
  );

  await expect(page.getByRole('alert')).toHaveText(
    'Some content of "Report.xlsx" is not supported and was not opened: ' +
      '1 chart, 1 pivot table, 1 picture, 1 formula that Gridworks cannot ' +
      'calculate (its last result was kept as a value).',
  );
  await expect(sheetTabs(page)).toHaveText(['Report']);
  await expectCell(page, 0, 1, 'B1 9');

  await page.getByRole('button', { name: 'Close' }).click();
  await expect(page.getByRole('alert')).toHaveCount(0);
});

test('a file that is not an .xlsx file leaves the workbook as it was', async ({
  page,
}) => {
  await openApp(page);
  await clickCell(page, 0, 0);
  await enter(page, 'kept');
  await openFile(page, 'notes.xlsx', new TextEncoder().encode('plain text'));

  await expect(page.getByRole('alert')).toHaveText(
    '"notes.xlsx" could not be opened because it is not a valid Excel ' +
      'workbook (.xlsx).',
  );
  await expect(sheetTabs(page)).toHaveText(['Sheet1']);
  await expectCell(page, 0, 0, 'A1 kept');
});

test.describe('in Chinese and the dark theme', () => {
  test.use({ locale: 'zh-CN', colorScheme: 'dark' });

  test('the File tab and the message bar are translated and dark', async ({
    page,
  }) => {
    await openApp(page);
    await page.getByRole('tab', { name: '文件' }).click();
    await page.getByRole('button', { name: '打开' }).hover();
    await expect(
      page.getByRole('tooltip', {
        name: '从计算机中打开 Excel 工作簿 (.xlsx)',
      }),
    ).toBeVisible();
    const chooser = page.waitForEvent('filechooser');
    await page.getByRole('button', { name: '打开' }).click();
    await (
      await chooser
    ).setFiles({
      name: 'notes.xlsx',
      mimeType: XLSX_TYPE,
      buffer: Buffer.from('plain text'),
    });

    const alert = page.getByRole('alert');
    await expect(alert).toHaveText(
      '无法打开“notes.xlsx”，因为它不是有效的 Excel 工作簿 (.xlsx)。',
    );
    const background = await alert.evaluate(
      (node) => getComputedStyle(node).backgroundColor,
    );
    const [red, green, blue] = background.match(/\d+/g)!.map(Number);
    expect((red! + green! + blue!) / 3).toBeLessThan(128);
    await page.getByRole('button', { name: '关闭' }).click();
    await expect(alert).toHaveCount(0);
  });
});
