import type { ErrorValue, Value } from '../value';
import { areaOf, evaluate, SHEET_COLUMNS, SHEET_ROWS } from './evaluate';
import { LIBRARY } from './library';
import { error, type Area, type ErrorCode } from './values';

export { evaluate, SHEET_COLUMNS, SHEET_ROWS };
export type { Area };

/**
 * One end of a range. A whole column has no row, and a whole row no column.
 */
export interface Corner {
  readonly row: number | null;
  readonly col: number | null;
  readonly rowAbsolute: boolean;
  readonly colAbsolute: boolean;
}

export type BinaryOperator =
  | '+'
  | '-'
  | '*'
  | '/'
  | '^'
  | '&'
  | '='
  | '<>'
  | '<'
  | '>'
  | '<='
  | '>='
  /** The intersection of two ranges, written as a space. */
  | ' ';

/**
 * A parsed formula. A reference without `sheet` is to the formula's own
 * sheet; `sheet` is the name as written.
 */
export type Ast =
  | { readonly type: 'number'; readonly value: number }
  | { readonly type: 'string'; readonly value: string }
  | { readonly type: 'boolean'; readonly value: boolean }
  | { readonly type: 'error'; readonly value: ErrorValue }
  | {
      readonly type: 'range';
      readonly sheet?: string;
      readonly start: Corner;
      /** Absent for a single cell. */
      readonly end?: Corner;
    }
  | {
      readonly type: 'unary';
      readonly operator: '-' | '+';
      readonly operand: Ast;
    }
  | { readonly type: 'percent'; readonly operand: Ast }
  | {
      readonly type: 'binary';
      readonly operator: BinaryOperator;
      readonly left: Ast;
      readonly right: Ast;
    }
  | { readonly type: 'paren'; readonly inner: Ast }
  | {
      readonly type: 'call';
      readonly name: string;
      readonly args: readonly Ast[];
    }
  /** An argument left out, as the second one of `IF(A1,,2)`. */
  | { readonly type: 'missing' }
  /** A name that is not a function or a reference, as written. */
  | { readonly type: 'name'; readonly name: string };

/** What a formula can read while it is evaluated. */
export interface Context {
  /** Where the formula is, for a range used where one value is expected. */
  readonly row: number;
  readonly col: number;
  /**
   * The value of a cell of the sheet named, or of the formula's own sheet
   * when `sheet` is absent; `#REF!` when no sheet has that name.
   */
  cell(sheet: string | undefined, row: number, col: number): Value;
  /**
   * Visits the cells of an area that are not empty, in no set order, with
   * their values and places; gives `#REF!` when no sheet has that name.
   */
  cells(
    sheet: string | undefined,
    area: Area,
    visit: (value: Value, row: number, col: number) => void,
  ): ErrorValue | undefined;
}

/**
 * A change to the workbook that references follow: a sheet renamed, or the
 * formula moved by rows and columns, which moves its relative references
 * as Excel's shared formulas and copies do.
 */
export type Move =
  | {
      readonly type: 'renameSheet';
      readonly from: string;
      readonly to: string;
    }
  | { readonly type: 'offset'; readonly rows: number; readonly cols: number };

// ---------------------------------------------------------------------------
// Parsing

/** How tightly each operator binds; unary minus binds tighter than all. */
const PRECEDENCE: Record<BinaryOperator | '%', number> = {
  '=': 1,
  '<>': 1,
  '<': 1,
  '>': 1,
  '<=': 1,
  '>=': 1,
  '&': 2,
  '+': 3,
  '-': 3,
  '*': 4,
  '/': 4,
  '^': 5,
  '%': 6,
  ' ': 8,
};
const UNARY = 7;

const OPERATOR = /<>|<=|>=|[-+*/^&=<>%]/y;
const NUMBER = /(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/y;
const CELL = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/y;
const COLUMN = /(\$?)([A-Za-z]{1,3})/y;
const ROW = /(\$?)(\d+)/y;
const NAME = /[A-Za-z_¡-￿][\w.¡-￿]*/y;
const QUOTED_SHEET = /'((?:[^']|'')+)'!/y;
const ERROR = /#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)/iy;
/** What may not follow a reference: more of a name, or a call. */
const NAME_GOES_ON = /[\w.(!¡-￿]/y;

const MAX_ROW = 1_048_576;
const MAX_COLUMN = 16_384;

/** The prefix that files give the names of functions newer than 2007. */
const PREFIX = '_XLFN.';

class Parser {
  private at = 0;

  constructor(private readonly text: string) {}

  /** The whole text as one expression, or `null`. */
  formula(): Ast | null {
    const ast = this.expression(0);
    this.skipSpace();
    return ast && this.at === this.text.length ? ast : null;
  }

  private expression(minimum: number): Ast | null {
    let left = this.prefix();
    while (left) {
      const operator = this.operator(left);
      if (!operator || PRECEDENCE[operator] < minimum) break;
      this.at += operator === ' ' ? 0 : operator.length;
      if (operator === '%') {
        left = { type: 'percent', operand: left };
        continue;
      }
      // Every binary operator of Excel groups from the left, `^` too.
      const right = this.expression(PRECEDENCE[operator] + 1);
      if (!right) return null;
      left = { type: 'binary', operator, left, right };
    }
    return left;
  }

  /**
   * The operator that follows, without taking it. Spaces between two
   * references are the intersection operator.
   */
  private operator(left: Ast): BinaryOperator | '%' | null {
    const before = this.at;
    this.skipSpace();
    const spaced = this.at > before;
    const match = this.match(OPERATOR);
    if (match) return match[0] as BinaryOperator | '%';
    const intersection =
      spaced &&
      isReference(left) &&
      /[A-Za-z$'(]/.test(this.text[this.at] ?? '');
    // The spaces are left for the next operand, or for an outer operator.
    this.at = before;
    return intersection ? ' ' : null;
  }

  private prefix(): Ast | null {
    this.skipSpace();
    const sign = this.text[this.at];
    if (sign === '-' || sign === '+') {
      this.at++;
      const operand = this.expression(UNARY);
      return operand && { type: 'unary', operator: sign, operand };
    }
    return this.operand();
  }

  private operand(): Ast | null {
    const { text } = this;
    const char = text[this.at];
    if (char === '(') {
      this.at++;
      const inner = this.expression(0);
      this.skipSpace();
      if (!inner || text[this.at] !== ')') return null;
      this.at++;
      return { type: 'paren', inner };
    }
    if (char === '"') return this.string();
    const errorMatch = this.take(ERROR);
    if (errorMatch) {
      const code = errorMatch[0].toUpperCase() as ErrorCode;
      return { type: 'error', value: error(code) };
    }
    const quoted = this.take(QUOTED_SHEET);
    if (quoted) return this.range(quoted[1]!.replaceAll("''", "'"));
    const reference = this.range(undefined);
    if (reference) return reference;
    const number = this.take(NUMBER);
    if (number) return { type: 'number', value: Number(number[0]) };
    const name = this.take(NAME);
    if (!name) return null;
    if (text[this.at] === '!') {
      this.at++;
      return this.range(name[0]);
    }
    if (text[this.at] === '(') return this.call(name[0]);
    const upper = name[0].toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') {
      return { type: 'boolean', value: upper === 'TRUE' };
    }
    return { type: 'name', name: name[0] };
  }

  private string(): Ast | null {
    const end = /"((?:[^"]|"")*)"/y;
    const match = this.take(end);
    return match && { type: 'string', value: match[1]!.replaceAll('""', '"') };
  }

  private call(name: string): Ast | null {
    this.at++; // The opening parenthesis.
    const args: Ast[] = [];
    this.skipSpace();
    if (this.text[this.at] === ')') {
      this.at++;
    } else {
      for (;;) {
        this.skipSpace();
        const next = this.text[this.at];
        const arg =
          next === ',' || next === ')'
            ? ({ type: 'missing' } as const)
            : this.expression(0);
        if (!arg) return null;
        args.push(arg);
        this.skipSpace();
        const separator = this.text[this.at++];
        if (separator === ')') break;
        if (separator !== ',') return null;
      }
    }
    let upper = name.toUpperCase();
    if (upper.startsWith(PREFIX) && upper.slice(PREFIX.length) in LIBRARY) {
      upper = upper.slice(PREFIX.length);
    }
    const spec = LIBRARY[upper];
    if (!spec) return { type: 'call', name, args };
    const { length } = args;
    if (length < spec.min || length > spec.max) return null;
    if (spec.pairs && (length - spec.min) % 2 !== 0) return null;
    return { type: 'call', name: upper, args };
  }

  /**
   * A cell, a range of cells, a whole column or a whole row; or `null`,
   * having taken nothing when there is no sheet.
   */
  private range(sheet: string | undefined): Ast | null {
    const start = this.at;
    const found =
      this.cellRange() ??
      this.lineRange(COLUMN, 'col') ??
      this.lineRange(ROW, 'row');
    if (!found || this.match(NAME_GOES_ON)) {
      this.at = start;
      return null;
    }
    return { type: 'range', ...(sheet !== undefined && { sheet }), ...found };
  }

  private cellRange(): { start: Corner; end?: Corner } | null {
    const start = this.corner();
    if (!start) return null;
    const before = this.at;
    if (this.text[this.at] === ':') {
      this.at++;
      const end = this.corner();
      if (end) return { start, end };
    }
    this.at = before;
    return { start };
  }

  private corner(): Corner | null {
    const match = this.match(CELL);
    if (!match) return null;
    const [text, colDollar, letters, rowDollar, digits] = match;
    const col = columnNumber(letters!);
    const row = Number(digits);
    if (col > MAX_COLUMN || row < 1 || row > MAX_ROW) return null;
    this.at += text.length;
    return {
      row: row - 1,
      col: col - 1,
      rowAbsolute: rowDollar === '$',
      colAbsolute: colDollar === '$',
    };
  }

  /** `A:C` or `1:3`, each end with an optional `$`. */
  private lineRange(
    pattern: RegExp,
    kind: 'row' | 'col',
  ): { start: Corner; end: Corner } | null {
    const before = this.at;
    const ends: Corner[] = [];
    for (let side = 0; side < 2; side++) {
      const match = this.take(pattern);
      if (!match || (side === 0 && this.text[this.at++] !== ':')) break;
      const [, dollar, label] = match;
      const number = kind === 'col' ? columnNumber(label!) : Number(label);
      if (number < 1 || number > (kind === 'col' ? MAX_COLUMN : MAX_ROW)) break;
      const absolute = dollar === '$';
      ends.push(
        kind === 'col'
          ? {
              row: null,
              col: number - 1,
              rowAbsolute: false,
              colAbsolute: absolute,
            }
          : {
              row: number - 1,
              col: null,
              rowAbsolute: absolute,
              colAbsolute: false,
            },
      );
    }
    if (ends.length < 2) {
      this.at = before;
      return null;
    }
    return { start: ends[0]!, end: ends[1]! };
  }

  private skipSpace(): void {
    while (this.text[this.at] === ' ') this.at++;
  }

  /** Matches a sticky pattern here, without taking it. */
  private match(pattern: RegExp): RegExpExecArray | null {
    pattern.lastIndex = this.at;
    return pattern.exec(this.text);
  }

  /** Matches a sticky pattern here, and takes it. */
  private take(pattern: RegExp): RegExpExecArray | null {
    const match = this.match(pattern);
    if (match) this.at += match[0].length;
    return match;
  }
}

function isReference(ast: Ast): boolean {
  return (
    ast.type === 'range' ||
    ast.type === 'paren' ||
    (ast.type === 'binary' && ast.operator === ' ')
  );
}

function columnNumber(letters: string): number {
  let col = 0;
  for (const letter of letters.toUpperCase()) {
    col = col * 26 + letter.charCodeAt(0) - 64;
  }
  return col;
}

/**
 * Parses the text of a formula, without its leading `=`. Returns `null`
 * when the text is not a valid formula.
 */
export function parse(text: string): Ast | null {
  return new Parser(text).formula();
}

// ---------------------------------------------------------------------------
// Printing and rewriting

/**
 * The text of a formula, without its leading `=`, as Excel writes it; for
 * a `file`, with the prefix that files give newer functions.
 */
export function print(ast: Ast, { file = false } = {}): string {
  const print_ = (node: Ast) => print(node, { file });
  switch (ast.type) {
    case 'number':
      return String(ast.value).toUpperCase();
    case 'string':
      return `"${ast.value.replaceAll('"', '""')}"`;
    case 'boolean':
      return ast.value ? 'TRUE' : 'FALSE';
    case 'error':
      return ast.value.error;
    case 'range': {
      const prefix =
        ast.sheet === undefined ? '' : `${quoteSheetName(ast.sheet)}!`;
      const end = ast.end ? `:${corner(ast.end)}` : '';
      return prefix + corner(ast.start) + end;
    }
    case 'unary':
      return ast.operator + print_(ast.operand);
    case 'percent':
      return `${print_(ast.operand)}%`;
    case 'binary':
      return print_(ast.left) + ast.operator + print_(ast.right);
    case 'paren':
      return `(${print_(ast.inner)})`;
    case 'call': {
      const prefix = file && LIBRARY[ast.name]?.prefixed ? '_xlfn.' : '';
      return `${prefix}${ast.name}(${ast.args.map(print_).join(',')})`;
    }
    case 'missing':
      return '';
    case 'name':
      return ast.name;
  }
}

function corner({ row, col, rowAbsolute, colAbsolute }: Corner): string {
  const column = col === null ? '' : (colAbsolute ? '$' : '') + columnName(col);
  const line = row === null ? '' : (rowAbsolute ? '$' : '') + (row + 1);
  return column + line;
}

/** A sheet name as a reference writes it: quoted unless it is a plain name. */
export function quoteSheetName(name: string): string {
  const plain =
    /^[A-Za-z_][\w.]*$/.test(name) &&
    !/^[A-Za-z]{1,3}\d+$/.test(name) &&
    !/^(TRUE|FALSE)$/i.test(name);
  return plain ? name : `'${name.replaceAll("'", "''")}'`;
}

/**
 * The formula with its references following a move, or the same formula
 * when no reference changes.
 */
export function shift(ast: Ast, move: Move): Ast {
  const range =
    move.type === 'offset' ? offset(move) : renamed(move.from, move.to);
  const visit = (node: Ast): Ast => {
    switch (node.type) {
      case 'range':
        return range(node);
      case 'unary':
      case 'percent': {
        const operand = visit(node.operand);
        return operand === node.operand ? node : { ...node, operand };
      }
      case 'paren': {
        const inner = visit(node.inner);
        return inner === node.inner ? node : { ...node, inner };
      }
      case 'binary': {
        const left = visit(node.left);
        const right = visit(node.right);
        return left === node.left && right === node.right
          ? node
          : { ...node, left, right };
      }
      case 'call': {
        const args = node.args.map(visit);
        return args.every((arg, index) => arg === node.args[index])
          ? node
          : { ...node, args };
      }
      default:
        return node;
    }
  };
  return visit(ast);
}

type RangeNode = Extract<Ast, { type: 'range' }>;

function renamed(from: string, to: string): (node: RangeNode) => Ast {
  const folded = from.toLowerCase();
  return (node) =>
    node.sheet?.toLowerCase() === folded ? { ...node, sheet: to } : node;
}

/** Moves relative references; one moved off the sheet becomes `#REF!`. */
function offset({ rows, cols }: { rows: number; cols: number }) {
  const move = (corner: Corner): Corner | null => {
    const row =
      corner.row === null || corner.rowAbsolute
        ? corner.row
        : corner.row + rows;
    const col =
      corner.col === null || corner.colAbsolute
        ? corner.col
        : corner.col + cols;
    const inside =
      (row === null || (row >= 0 && row < SHEET_ROWS)) &&
      (col === null || (col >= 0 && col < SHEET_COLUMNS));
    return inside ? { ...corner, row, col } : null;
  };
  return (node: RangeNode): Ast => {
    const start = move(node.start);
    const end = node.end && move(node.end);
    if (!start || end === null) return { type: 'error', value: error('#REF!') };
    return { ...node, start, ...(end && { end }) };
  };
}

/**
 * Whether Gridworks can calculate a formula: whether every function it
 * calls is in the library and it uses no other name.
 */
export function calculable(ast: Ast): boolean {
  switch (ast.type) {
    case 'name':
      return false;
    case 'call':
      return ast.name in LIBRARY && ast.args.every(calculable);
    case 'unary':
    case 'percent':
      return calculable(ast.operand);
    case 'paren':
      return calculable(ast.inner);
    case 'binary':
      return calculable(ast.left) && calculable(ast.right);
    default:
      return true;
  }
}

/** The areas a formula reads, with the sheet named for each, in order. */
export function references(
  ast: Ast,
): { readonly sheet?: string; readonly area: Area }[] {
  const found: { sheet?: string; area: Area }[] = [];
  const visit = (node: Ast): void => {
    switch (node.type) {
      case 'range':
        found.push({
          ...(node.sheet !== undefined && { sheet: node.sheet }),
          area: areaOf(node),
        });
        return;
      case 'unary':
      case 'percent':
        return visit(node.operand);
      case 'paren':
        return visit(node.inner);
      case 'binary':
        visit(node.left);
        return visit(node.right);
      case 'call':
        return node.args.forEach(visit);
      default:
        return;
    }
  };
  visit(ast);
  return found;
}

/** The letters of a column, counted from zero: `A`, ..., `Z`, `AA`, ... */
export function columnName(col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}
