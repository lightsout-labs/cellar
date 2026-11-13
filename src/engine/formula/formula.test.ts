import { describe, expect, it } from 'vitest';
import type { Value } from '../value';
import { evaluate, parse } from './index';

function calculate(text: string, cells: Record<string, Value> = {}): Value {
  const ast = parse(text);
  if (!ast) throw new Error(`not a formula: ${text}`);
  return evaluate(ast, {
    cell: (row, col) => cells[`${row},${col}`] ?? null,
  });
}

describe('formula', () => {
  it('adds a referenced cell and a number', () => {
    expect(calculate('A1+1', { '0,0': 2 })).toBe(3);
  });

  it('adds several operands, ignoring spaces and letter case', () => {
    expect(calculate(' 1.5 + b3 + 2 ', { '2,1': 10 })).toBe(13.5);
  });

  it('reads columns beyond Z', () => {
    expect(calculate('AB2', { '1,27': 7 })).toBe(7);
  });

  it('treats an empty cell as zero', () => {
    expect(calculate('A1+1')).toBe(1);
  });

  it('gives #VALUE! when an operand is text', () => {
    expect(calculate('A1+1', { '0,0': 'abc' })).toEqual({ error: '#VALUE!' });
  });

  it('passes on an error from a referenced cell', () => {
    const error = { error: '#VALUE!' };
    expect(calculate('1+A1', { '0,0': error })).toEqual(error);
  });

  it.each(['', 'A1+', '+1', '1 2', 'A0', 'hello', '1+*2'])(
    'rejects %j',
    (text) => {
      expect(parse(text)).toBeNull();
    },
  );
});
