import { describe, expect, it } from 'vitest';
import type { Value } from '../value';
import {
  evaluate,
  parse,
  print,
  references,
  shift,
  SHEET_COLUMNS,
  SHEET_ROWS,
  type Area,
  type Context,
} from './index';

/** The sheets of the test workbook, besides the formula's own. */
const SHEETS = ['Sheet2', 'My Sheet', "It's"];

/**
 * A context over cells named as in `A1` or `Sheet2!A1`, for a formula in
 * the cell `at`.
 */
function context(cells: Record<string, Value>, at = 'Z100'): Context {
  const position = locate(at);
  const sheetOf = (name: string | undefined) =>
    name === undefined
      ? ''
      : SHEETS.find((each) => each.toLowerCase() === name.toLowerCase());
  const located = Object.entries(cells).map(([name, value]) => {
    const [sheet, cell] = name.includes('!') ? name.split('!') : ['', name];
    return { sheet: sheet!, ...locate(cell!), value };
  });
  return {
    ...position,
    cell(sheet, row, col) {
      const found = sheetOf(sheet);
      if (found === undefined) return { error: '#REF!' };
      const cell = located.find(
        (each) => each.sheet === found && each.row === row && each.col === col,
      );
      return cell?.value ?? null;
    },
    cells(sheet, area: Area) {
      const found = sheetOf(sheet);
      if (found === undefined) return { error: '#REF!' };
      return located
        .filter(
          ({ sheet, row, col, value }) =>
            sheet === found &&
            value !== null &&
            row >= area.top &&
            row <= area.bottom &&
            col >= area.left &&
            col <= area.right,
        )
        .map(({ value }) => value);
    },
  };
}

function locate(name: string): { row: number; col: number } {
  const [, letters, digits] = /^([A-Z]+)(\d+)$/.exec(name)!;
  let col = 0;
  for (const letter of letters!) col = col * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(digits) - 1, col: col - 1 };
}

function calculate(
  text: string,
  cells: Record<string, Value> = {},
  at?: string,
): Value {
  const ast = parse(text);
  if (!ast) throw new Error(`not a formula: ${text}`);
  return evaluate(ast, context(cells, at));
}

const error = (code: string) => ({ error: code });

describe('formula operators', () => {
  it.each<[string, Value]>([
    ['1+2*3', 7],
    ['(1+2)*3', 9],
    ['10-2-3', 5],
    ['8/2/2', 2],
    ['2^3^2', 64],
    ['-2^2', 4],
    ['2^-1', 0.5],
    ['50%', 0.5],
    ['2*50%^2', 0.5],
    ['--3', 3],
    ['+"a"', 'a'],
    ['1+2&"x"', '3x'],
    ['"a"&1+1', 'a2'],
    ['0.1+0.2&""', '0.3'],
    ['3&TRUE', '3TRUE'],
    ['1+1=2', true],
    ['0.1+0.2=0.3', true],
    ['1=1&""', false],
    ['1<2', true],
    ['2>=3', false],
    ['1<=1', true],
    ['2>1', true],
    ['"a"<>"A"', false],
    ['"b">"a"', true],
    ['"a"<1', false],
    ['TRUE>"z"', true],
    ['"3"+1', 4],
    ['TRUE+1', 2],
    [' 1.5 + 2 ', 3.5],
    ['1e3+.5', 1000.5],
  ])('%s gives %j, with Excel precedence and coercion', (text, expected) => {
    expect(calculate(text)).toEqual(expected);
  });

  it('treats an empty cell as zero in arithmetic and as nothing in text', () => {
    expect(calculate('A1+1')).toBe(1);
    expect(calculate('A1&"x"')).toBe('x');
    expect(calculate('A1=0')).toBe(true);
    expect(calculate('A1=""')).toBe(true);
    expect(calculate('A1')).toBe(0);
  });
});

describe('formula references', () => {
  it('reads relative, absolute and mixed references, in any case', () => {
    expect(calculate('$A$1+A$1+$A1+a1', { A1: 2 })).toBe(8);
    expect(calculate('AB2', { AB2: 7 })).toBe(7);
  });

  it('reads ranges, whole columns and whole rows', () => {
    const cells = { A1: 1, B2: 2, A100000: 4, CV1: 8 };
    expect(calculate('SUM(A1:B2)', cells)).toBe(3);
    expect(calculate('SUM($B$2:a1)', cells)).toBe(3);
    expect(calculate('SUM(A:A)', cells)).toBe(5);
    expect(calculate('SUM(1:1)', cells)).toBe(9);
    expect(calculate('SUM($A:$B,2:2)', cells)).toBe(9);
  });

  it('reads other sheets, by plain or quoted name, in any case', () => {
    const cells = { 'Sheet2!A1': 1, 'My Sheet!A1': 2, "It's!A1": 4 };
    expect(calculate("sheet2!A1+'My Sheet'!A1+'It''s'!A1", cells)).toBe(7);
    expect(calculate('SUM(Sheet2!A1:B2)', cells)).toBe(1);
    expect(calculate('A1', cells)).toBe(0);
  });

  it('gives the cell of a range in the row or column of the formula', () => {
    const cells = { A1: 1, A2: 2, A3: 3, C1: 4 };
    expect(calculate('A1:A3*10', cells, 'B2')).toBe(20);
    expect(calculate('A1:C1', cells, 'C5')).toBe(4);
    expect(calculate('A1:A3', cells, 'B5')).toEqual(error('#VALUE!'));
  });

  it('intersects ranges separated by a space', () => {
    const cells = { A1: 1, B1: 2, B2: 4, C2: 8 };
    expect(calculate('SUM(A1:B2 B1:C3)', cells)).toBe(6);
    expect(calculate('SUM(A1:A2 C1:C2)', cells)).toEqual(error('#NULL!'));
  });

  it('lists the areas a formula reads, with their sheets', () => {
    const ast = parse("A1+SUM('My Sheet'!B1:C2,A:A,2:2)")!;
    expect(references(ast)).toEqual([
      { area: { top: 0, left: 0, bottom: 0, right: 0 } },
      { sheet: 'My Sheet', area: { top: 0, left: 1, bottom: 1, right: 2 } },
      { area: { top: 0, left: 0, bottom: SHEET_ROWS - 1, right: 0 } },
      { area: { top: 1, left: 0, bottom: 1, right: SHEET_COLUMNS - 1 } },
    ]);
  });
});

describe('formula functions', () => {
  const cells = { A1: 1, A2: 'text', A3: true, A4: 5 };

  it('totals, averages and finds the extremes of numbers, skipping other cells of ranges', () => {
    expect(calculate('SUM(A1:A4)', cells)).toBe(6);
    expect(calculate('SUM(A1:A4,"2",TRUE,3)', cells)).toBe(12);
    expect(calculate('AVERAGE(A1:A4)', cells)).toBe(3);
    expect(calculate('AVERAGE(A1:A4,0)', cells)).toBe(2);
    expect(calculate('MIN(A1:A4)', cells)).toBe(1);
    expect(calculate('MAX(A1:A4,-1)', cells)).toBe(5);
    expect(calculate('MIN(A2:A3)', cells)).toBe(0);
    expect(calculate('max(A2)', cells)).toBe(0);
    expect(calculate('SUM(1,)', cells)).toBe(1);
  });

  it('counts numbers, and typed values that are numbers', () => {
    expect(calculate('COUNT(A1:A4)', cells)).toBe(2);
    expect(calculate('COUNT(A1:A4,"3","x",TRUE,1/0)', cells)).toBe(4);
    expect(calculate('COUNT(A1:A4)', { ...cells, A2: error('#N/A') })).toBe(2);
  });

  it('chooses a branch by a condition, evaluating only that branch', () => {
    expect(calculate('IF(A1>0,"yes","no")', cells)).toBe('yes');
    expect(calculate('IF(SUM(A1:A4)>10,1,2)', cells)).toBe(2);
    expect(calculate('IF(0,1/0,"safe")', cells)).toBe('safe');
    expect(calculate('IF(FALSE,1)', cells)).toBe(false);
    expect(calculate('IF(TRUE,)', cells)).toBe(0);
    expect(calculate('IF("true",1,2)', cells)).toBe(1);
    expect(calculate('SUM(IF(TRUE,A1:A4))', cells)).toBe(6);
    expect(calculate('IF("x",1,2)', cells)).toEqual(error('#VALUE!'));
  });
});

describe('formula errors', () => {
  it.each<[string, string]>([
    ['1/0', '#DIV/0!'],
    ['AVERAGE(A1:A2)', '#DIV/0!'],
    ['#N/A', '#N/A'],
    ['NOSUCH(1)', '#NAME?'],
    ['total', '#NAME?'],
    ['A0', '#NAME?'],
    ['SUM(A1:A2 B1:B2)', '#NULL!'],
    ['10^400', '#NUM!'],
    ['(-8)^(1/3)', '#NUM!'],
    ['0^0', '#NUM!'],
    ['0^-1', '#DIV/0!'],
    ['SUM(1E308,1E308)', '#NUM!'],
    ['Nowhere!A1', '#REF!'],
    ['SUM(Nowhere!A1:A2)', '#REF!'],
    ['"a"+1', '#VALUE!'],
    ['-"a"', '#VALUE!'],
    ['#div/0!', '#DIV/0!'],
  ])('%s gives %s', (text, code) => {
    expect(calculate(text)).toEqual(error(code));
  });

  it('passes errors on through operators and functions, the leftmost first', () => {
    const cells = { A1: error('#N/A'), A2: 1 };
    expect(calculate('1+A1', cells)).toEqual(error('#N/A'));
    expect(calculate('A1&"x"', cells)).toEqual(error('#N/A'));
    expect(calculate('A1=1', cells)).toEqual(error('#N/A'));
    expect(calculate('-A1%', cells)).toEqual(error('#N/A'));
    expect(calculate('SUM(A1:A2)', cells)).toEqual(error('#N/A'));
    expect(calculate('MAX(A2,A1)', cells)).toEqual(error('#N/A'));
    expect(calculate('IF(A1,1,2)', cells)).toEqual(error('#N/A'));
    expect(calculate('1/0+A1', cells)).toEqual(error('#DIV/0!'));
  });
});

describe('formula text', () => {
  it.each([
    '',
    'A1+',
    '+',
    '1 2',
    '(1',
    '1)',
    '1+*2',
    'SUM()',
    'IF(1)',
    'IF(1,2,3,4)',
    'SUM(1',
    'A1:',
    'A1:B',
    '"abc',
    'Sheet2!',
    "'Sheet2!A1",
    '#FOO!',
    'SUM(A1;A2)',
  ])('rejects %j', (text) => {
    expect(parse(text)).toBeNull();
  });

  it.each([
    [' sum( a1 , $B$2:c3 ) ', 'SUM(A1,$B$2:C3)'],
    ["'my sheet'!a1+sheet2!B$2", "'my sheet'!A1+sheet2!B$2"],
    ["'A1'!A1+'It''s'!A1", "'A1'!A1+'It''s'!A1"],
    ['Sheet2!a:$b+1:3', 'Sheet2!A:$B+1:3'],
    ['-2^2', '-2^2'],
    ['(1+2)%', '(1+2)%'],
    ['"a""b"&TRUE', '"a""b"&TRUE'],
    ['A1:B2 B1 B1:C1', 'A1:B2 B1 B1:C1'],
    ['IF(A1,,)', 'IF(A1,,)'],
    ['#n/a<>total', '#N/A<>total'],
    ['1.50+1E3', '1.5+1000'],
  ])('prints %j as %j', (text, printed) => {
    expect(print(parse(text)!)).toBe(printed);
  });

  it('renames a sheet in the references to it, in any case', () => {
    const ast = parse('Sheet2!A1+sheet2!B1:B2+A1+Other!A1')!;
    const move = {
      type: 'renameSheet',
      from: 'SHEET2',
      to: 'My Data',
    } as const;
    expect(print(shift(ast, move))).toBe(
      "'My Data'!A1+'My Data'!B1:B2+A1+Other!A1",
    );
    const untouched = parse('A1+Other!A1')!;
    expect(shift(untouched, move)).toBe(untouched);
  });
});
