import { describe, expect, it } from 'vitest';
import { droppedMessage, workbookName } from './files';

describe('workbookName', () => {
  it('is the name of the file without its extension', () => {
    expect(workbookName('Budget 2026.xlsx')).toBe('Budget 2026');
    expect(workbookName('report.final.XLSX')).toBe('report.final');
    expect(workbookName('notes')).toBe('notes');
  });
});

describe('droppedMessage', () => {
  it('names the file and counts each feature dropped, in either language', () => {
    const dropped = [
      { feature: 'chart', count: 2 },
      { feature: 'pivotTable', count: 1 },
      { feature: 'formula', count: 3 },
    ] as const;
    expect(droppedMessage('Sales.xlsx', dropped, 'en')).toBe(
      'Some content of "Sales.xlsx" is not supported and was not opened: ' +
        '2 charts, 1 pivot table, 3 formulas that Gridworks cannot calculate ' +
        '(their last results were kept as values).',
    );
    expect(droppedMessage('Sales.xlsx', dropped, 'zh')).toBe(
      '“Sales.xlsx”中有部分内容不受支持，未能打开：' +
        '2 个图表、1 个数据透视表、3 个 Gridworks 无法计算的公式（已将其最后结果保留为值）。',
    );
  });

  it('speaks of one picture and many cells', () => {
    expect(
      droppedMessage(
        'a.xlsx',
        [
          { feature: 'picture', count: 1 },
          { feature: 'cell', count: 1200 },
        ],
        'en',
      ),
    ).toBe(
      'Some content of "a.xlsx" is not supported and was not opened: ' +
        '1 picture, 1,200 cells beyond row 100,000 or column CV.',
    );
  });
});
