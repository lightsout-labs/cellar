import { describe, expect, it } from 'vitest';
import {
  complete,
  cycleReference,
  formulaHelp,
  insertReference,
  pickable,
} from './formula';

describe('function completion', () => {
  it('lists the functions that start with the name typed at the caret', () => {
    const { completion } = formulaHelp('=1+su', 5);
    expect(completion).toMatchObject({ start: 3, end: 5 });
    expect(completion!.names).toContain('SUM');
    expect(completion!.names).toContain('SUMIFS');
    expect(completion!.names.every((name) => name.startsWith('SU'))).toBe(true);
    // A name that looks like a cell, as LOG1, still completes.
    expect(formulaHelp('=LOG1', 5).completion?.names).toEqual(['LOG10']);
  });

  it('offers nothing where no function name is being typed', () => {
    for (const [text, caret] of [
      ['=XYZ', 4],
      ['=SUM(', 5],
      ['=SUM', 2], // The caret is inside the name.
      ['="su', 4],
      ['=A1', 3],
      ['=$A', 3],
      ['su', 2],
    ] as const) {
      expect([text, formulaHelp(text, caret).completion]).toEqual([text, null]);
    }
  });

  it('replaces the name typed with the function chosen and its parenthesis', () => {
    expect(complete('=1+su*2', 5, 'SUM')).toEqual({
      text: '=1+SUM(*2',
      caret: 7,
    });
  });
});

describe('argument hint', () => {
  it('shows the innermost function whose parentheses hold the caret', () => {
    const hint = (text: string) => formulaHelp(text, text.length).hint;
    expect(hint('=SUM(A1,')).toEqual({
      name: 'SUM',
      arguments: ['number1', '[number2]', '...'],
      current: 1,
    });
    expect(hint('=IF(A1,SUM(1),')).toMatchObject({ name: 'IF', current: 2 });
    expect(hint('=IF(A1,SUM(1,')).toMatchObject({ name: 'SUM', current: 1 });
    expect(hint('=IF(A1,"a,b"')).toMatchObject({ name: 'IF', current: 1 });
    expect(hint('=round((1+2')).toMatchObject({ name: 'ROUND', current: 0 });
    expect(hint('=SUM(1)')).toBeNull();
    expect(hint('=FOO(1')).toBeNull();
    expect(hint('SUM(1')).toBeNull();
    expect(formulaHelp('=SUM(1,2)', 7).hint).toMatchObject({ current: 1 });
  });
});

describe('references', () => {
  it('gives each different reference its own color, and its area and sheet', () => {
    const { references } = formulaHelp("=A1+B2:C3+a1+'My sheet'!A1", 0);
    expect(references).toEqual([
      { start: 1, end: 3, color: 0, area: area(0, 0, 0, 0) },
      { start: 4, end: 9, color: 1, area: area(1, 1, 2, 2) },
      { start: 10, end: 12, color: 0, area: area(0, 0, 0, 0) },
      {
        start: 13,
        end: 26,
        color: 2,
        sheet: 'My sheet',
        area: area(0, 0, 0, 0),
      },
    ]);
    expect(formulaHelp('A1', 0).references).toEqual([]);
  });

  it('can take a reference where an operand starts, or replace the one just before', () => {
    expect(pickable('=', 1)).toBe(true);
    expect(pickable('=A1+ ', 5)).toBe(true);
    expect(pickable('=SUM(', 5)).toBe(true);
    expect(pickable('=SUM(A1,', 8)).toBe(true);
    expect(pickable('=A1:', 4)).toBe(true);
    expect(pickable('=A1', 3)).toBe(true);
    expect(pickable('=A1', 2)).toBe(false);
    expect(pickable('=5', 2)).toBe(false);
    expect(pickable('=SUM(1)', 7)).toBe(false);
    expect(pickable('=SUM', 4)).toBe(false);
    expect(pickable('5+', 2)).toBe(false);
  });

  it('inserts a reference at the caret, or replaces the one just before it', () => {
    expect(insertReference('=SUM()', 5, 'B2')).toEqual({
      text: '=SUM(B2)',
      caret: 7,
    });
    expect(insertReference('=SUM(B2)', 7, 'B2:C4')).toEqual({
      text: '=SUM(B2:C4)',
      caret: 10,
    });
    expect(insertReference('=1+Sheet2!A1', 12, 'Sheet2!B1')).toEqual({
      text: '=1+Sheet2!B1',
      caret: 12,
    });
  });

  it('cycles the reference at the caret through absolute and relative forms', () => {
    let edit = { text: '=A1+B2', caret: 2 };
    const seen: string[] = [];
    for (let step = 0; step < 4; step++) {
      edit = cycleReference(edit.text, edit.caret)!;
      seen.push(edit.text);
    }
    expect(seen).toEqual(['=$A$1+B2', '=A$1+B2', '=$A1+B2', '=A1+B2']);
    expect(edit.caret).toBe(3);
    expect(cycleReference('=SUM(A1:B2)', 10)).toEqual({
      text: '=SUM($A$1:$B$2)',
      caret: 14,
    });
    expect(cycleReference("='My sheet'!a1", 14)).toEqual({
      text: "='My sheet'!$A$1",
      caret: 16,
    });
    expect(cycleReference('=1+2', 2)).toBeNull();
    expect(cycleReference('A1', 2)).toBeNull();
  });
});

function area(top: number, left: number, bottom: number, right: number) {
  return { top, left, bottom, right };
}
