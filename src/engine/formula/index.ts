import { isError, VALUE_ERROR, type Value } from '../value';

/** A parsed formula. Rows and columns count from zero. */
export type Ast =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'ref'; readonly row: number; readonly col: number }
  | { readonly type: 'add'; readonly left: Ast; readonly right: Ast };

/** What a formula can read while it is evaluated. */
export interface Context {
  cell(row: number, col: number): Value;
}

const TOKEN = /\s*(?:(\d+(?:\.\d+)?)|([A-Za-z]+)(\d+)|(\+))/y;

/**
 * Parses the text of a formula, without its leading `=`. Returns `null`
 * when the text is not a valid formula.
 */
export function parse(text: string): Ast | null {
  TOKEN.lastIndex = 0;
  let ast: Ast | null = null;
  let expectOperand = true;
  while (text.slice(TOKEN.lastIndex).trim() !== '') {
    const match = TOKEN.exec(text);
    if (!match) return null;
    const [, number, letters, digits, plus] = match;
    if (Boolean(plus) === expectOperand) return null;
    if (!plus) {
      const operand = number
        ? ({ type: 'number', value: Number(number) } as const)
        : reference(letters!, digits!);
      if (!operand) return null;
      ast = ast ? { type: 'add', left: ast, right: operand } : operand;
    }
    expectOperand = !expectOperand;
  }
  return expectOperand ? null : ast;
}

function reference(letters: string, digits: string): Ast | null {
  const row = Number(digits) - 1;
  if (row < 0) return null;
  let col = 0;
  for (const letter of letters.toUpperCase()) {
    col = col * 26 + letter.charCodeAt(0) - 64;
  }
  return { type: 'ref', row, col: col - 1 };
}

/** Calculates the value of a parsed formula. */
export function evaluate(ast: Ast, context: Context): Value {
  switch (ast.type) {
    case 'number':
      return ast.value;
    case 'ref':
      return context.cell(ast.row, ast.col);
    case 'add': {
      const left = toNumber(evaluate(ast.left, context));
      if (isError(left)) return left;
      const right = toNumber(evaluate(ast.right, context));
      if (isError(right)) return right;
      return left + right;
    }
  }
}

/**
 * The number an operand stands for: an empty cell counts as zero, and
 * `TRUE` as one.
 */
function toNumber(value: Value): number | Extract<Value, object> {
  if (value === null) return 0;
  if (typeof value === 'string') return VALUE_ERROR;
  if (typeof value === 'boolean') return Number(value);
  return value;
}

/** The letters of a column, counted from zero: `A`, ..., `Z`, `AA`, ... */
export function columnName(col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}
