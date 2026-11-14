import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { UNSUPPORTED_PARTS, xlsxFile } from '../../../tests/xlsx';
import { Workbook, type Snapshot } from '../workbook';
import { readXlsx, writeXlsx } from '.';

/** The value, formula and display text of a cell given by its name. */
function cellOf(workbook: Workbook, sheet: string, name: string) {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(name)!;
  const col = letters!.charCodeAt(0) - 65;
  const { value, formula, display } = workbook.cell(
    sheet,
    Number(digits) - 1,
    col,
  );
  return { value, formula, display };
}

async function open(bytes: Uint8Array) {
  const { snapshot, dropped } = await readXlsx(bytes, 'Budget');
  return { workbook: Workbook.fromSnapshot(snapshot), snapshot, dropped };
}

describe('readXlsx', () => {
  it('reads the sheets of a file in order, with their names and tab colors', async () => {
    const { workbook, snapshot } = await open(
      await xlsxFile({
        sheets: [
          { name: 'Income', rows: '' },
          {
            name: 'Q&A 2026',
            before: '<sheetPr><tabColor rgb="FFFF0000"/></sheetPr>',
            rows: '',
          },
        ],
      }),
    );
    expect(snapshot.name).toBe('Budget');
    expect(workbook.sheets.map(({ name, color }) => ({ name, color }))).toEqual(
      [
        { name: 'Income', color: undefined },
        { name: 'Q&A 2026', color: '#ff0000' },
      ],
    );
  });

  it('reads numbers, shared and inline strings, booleans and errors', async () => {
    const { workbook, snapshot } = await open(
      await xlsxFile({
        strings: [
          '<t>Rent &amp; bills</t>',
          '<r><rPr><b/></rPr><t>Bold</t></r><r><t xml:space="preserve"> part</t></r>',
          '<t>Line_x000D_</t><rPh><t>ignored</t></rPh>',
        ],
        sheets: [
          {
            name: 'Sheet1',
            rows:
              '<row r="1"><c r="A1"><v>12.5</v></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>' +
              '<row r="3"><c r="A3" t="b"><v>1</v></c><c r="B3" t="inlineStr"><is><t>&lt;inline&gt;</t></is></c><c r="C3" t="e"><v>#N/A</v></c><c r="D3" t="s"><v>2</v></c></row>',
          },
        ],
      }),
    );
    const id = snapshot.sheets[0]!.id;
    expect(cellOf(workbook, id, 'A1').value).toBe(12.5);
    expect(cellOf(workbook, id, 'B1').value).toBe('Rent & bills');
    expect(cellOf(workbook, id, 'C1').value).toBe('Bold part');
    expect(cellOf(workbook, id, 'A3').value).toBe(true);
    expect(cellOf(workbook, id, 'B3').value).toBe('<inline>');
    expect(cellOf(workbook, id, 'C3').display).toBe('#N/A');
    expect(cellOf(workbook, id, 'D3').value).toBe('Line\r');
    expect(cellOf(workbook, id, 'A2').value).toBeNull();
  });

  it('reads number formats, built in and custom', async () => {
    const { workbook, snapshot } = await open(
      await xlsxFile({
        numberFormats: [[164, '0.0%']],
        cellFormats: [0, 14, 164],
        sheets: [
          {
            name: 'Sheet1',
            rows: '<row r="1"><c r="A1" s="1"><v>46340</v></c><c r="B1" s="2"><v>0.125</v></c></row>',
          },
        ],
      }),
    );
    const id = snapshot.sheets[0]!.id;
    expect(cellOf(workbook, id, 'A1').display).toBe('11/14/2026');
    expect(cellOf(workbook, id, 'B1').display).toBe('12.5%');
  });

  it('reads formulas, filling in those shared down a column, and calculates them', async () => {
    const { workbook, snapshot, dropped } = await open(
      await xlsxFile({
        sheets: [
          { name: 'Rates', rows: '<row r="1"><c r="A1"><v>2</v></c></row>' },
          {
            name: 'Sheet2',
            rows:
              '<row r="1"><c r="A1"><v>10</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*Rates!$A$1</f><v>20</v></c></row>' +
              '<row r="2"><c r="A2"><v>20</v></c><c r="B2"><f t="shared" si="0"/><v>40</v></c></row>' +
              '<row r="3"><c r="A3"><v>30</v></c><c r="B3"><f t="shared" si="0"/><v>60</v></c><c r="C3"><f>SUM(B1:B3)</f><v>120</v></c></row>',
          },
        ],
      }),
    );
    const id = snapshot.sheets[1]!.id;
    expect(cellOf(workbook, id, 'B3')).toEqual({
      value: 60,
      formula: '=A3*Rates!$A$1',
      display: '60',
    });
    expect(cellOf(workbook, id, 'C3')).toMatchObject({
      value: 120,
      formula: '=SUM(B1:B3)',
    });
    expect(dropped).toEqual([]);
  });

  it('opens a file with charts, pivot tables, pictures and unknown formulas, and lists what it dropped', async () => {
    const { workbook, snapshot, dropped } = await open(
      await xlsxFile({
        strings: ['<t>old text</t>'],
        sheets: [
          {
            name: 'Sheet1',
            rows:
              '<row r="1"><c r="A1"><v>3</v></c><c r="B1"><f>_xlfn.XLOOKUP(A1,C1:C2,D1:D2)</f><v>9</v></c><c r="C1" t="str"><f>[1]Other!A1</f><v>linked</v></c></row>' +
              '<row r="2"><c r="A2"><f>NOSUCHFUNCTION(A1)</f><v>#NAME?</v></c><c r="DA2"><v>1</v></c></row>' +
              '<row r="100001"><c r="A100001"><v>1</v></c></row>',
          },
        ],
        parts: UNSUPPORTED_PARTS,
      }),
    );
    const id = snapshot.sheets[0]!.id;
    expect(cellOf(workbook, id, 'A1').value).toBe(3);
    // The last result is kept in place of a formula that cannot be calculated.
    expect(cellOf(workbook, id, 'B1')).toMatchObject({
      value: 9,
      formula: null,
    });
    expect(cellOf(workbook, id, 'C1').value).toBe('linked');
    expect(dropped).toEqual([
      { feature: 'chart', count: 1 },
      { feature: 'pivotTable', count: 1 },
      { feature: 'picture', count: 1 },
      { feature: 'formula', count: 3 },
      { feature: 'cell', count: 2 },
    ]);
  });

  it('rejects a file that is not an .xlsx file', async () => {
    await expect(
      readXlsx(new TextEncoder().encode('just text'), 'notes'),
    ).rejects.toThrow();
    const zip = new JSZip();
    zip.file('readme.txt', 'not a workbook');
    await expect(
      readXlsx(await zip.generateAsync({ type: 'uint8array' }), 'archive'),
    ).rejects.toThrow();
  });

  it('opens a 10 MB file within 5 seconds (R47)', async () => {
    const bytes = await largeFile();
    expect(bytes.byteLength).toBeGreaterThan(10_000_000);
    const start = performance.now();
    const { workbook, snapshot } = await open(bytes);
    const id = snapshot.sheets[0]!.id;
    expect(cellOf(workbook, id, 'K2').value).toBe('text 2-02.5');
    expect(performance.now() - start).toBeLessThan(5000);
  }, 30_000);
});

describe('writeXlsx', () => {
  it('writes a file that reads back as the same sheets, values, formulas and formats', async () => {
    const snapshot: Snapshot = {
      version: 1,
      name: 'Round trip',
      styles: [{}, { numberFormat: 'm/d/yyyy' }, { numberFormat: '0.0%' }],
      sheets: [
        {
          id: 'sheet1',
          name: 'Data',
          color: '#00aa55',
          columns: [
            [1, 'two', true, 46340, null],
            [null, 0.5, 'Q&A <b>'],
          ],
          formulas: [
            [0, 1, '=A1+Totals!A1'],
            [3, 1, '=A1&" items"'],
          ],
          styleIds: [
            [0, 0, 0, 1],
            [0, 2],
          ],
        },
        {
          id: 'sheet2',
          name: 'Totals',
          columns: [[5]],
          formulas: [],
          styleIds: [],
        },
      ],
    };
    const before = Workbook.fromSnapshot(snapshot);
    const { snapshot: read, dropped } = await readXlsx(
      await writeXlsx(snapshot),
      'Round trip',
    );
    const after = Workbook.fromSnapshot(read);
    expect(dropped).toEqual([]);
    expect(after.sheets).toEqual(before.sheets);
    for (const sheet of before.sheets) {
      for (let row = 0; row < 5; row++) {
        for (let col = 0; col < 2; col++) {
          expect(after.cell(sheet.id, row, col)).toEqual(
            before.cell(sheet.id, row, col),
          );
        }
      }
    }
  });
});

/**
 * A file of more than 10 MB, as Excel would write a large table: 150,000
 * rows of ten numbers and texts and a formula.
 */
async function largeFile(): Promise<Uint8Array> {
  const strings: string[] = [];
  const rows: string[] = [];
  for (let r = 1; r <= 150_000; r++) {
    const cells: string[] = [];
    for (let c = 0; c < 10; c++) {
      const ref = String.fromCharCode(65 + c) + r;
      if (c % 3 === 0) {
        cells.push(`<c r="${ref}" t="s"><v>${strings.length}</v></c>`);
        strings.push(`<t>text ${r}-${c}</t>`);
      } else {
        cells.push(`<c r="${ref}"><v>${r * c + 0.5}</v></c>`);
      }
    }
    cells.push(
      `<c r="K${r}" t="str"><f>A${r}&amp;B${r}</f><v>text ${r}-0${r + 0.5}</v></c>`,
    );
    rows.push(`<row r="${r}">${cells.join('')}</row>`);
  }
  return xlsxFile({
    strings,
    sheets: [{ name: 'Data', rows: rows.join('') }],
  });
}
