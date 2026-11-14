import { parseInput } from '../format';
import { isError, type ErrorValue, type Value } from '../value';

/** The number of rows and columns of every sheet (R1). */
export const SHEET_ROWS = 100_000;
export const SHEET_COLUMNS = 100;

/** A rectangle of cells, its bounds included. Rows and columns count from zero. */
export interface Area {
  readonly top: number;
  readonly left: number;
  readonly bottom: number;
  readonly right: number;
}

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
  /** The values of the cells of an area that are not empty, or `#REF!`. */
  cells(sheet: string | undefined, area: Area): readonly Value[] | ErrorValue;
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

/** The seven error values of Excel (R13). */
type ErrorCode =
  '#NULL!' | '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#NUM!' | '#N/A';

const error = (code: ErrorCode): ErrorValue => ({ error: code });
const VALUE = error('#VALUE!');
const DIV0 = error('#DIV/0!');
const NUM = error('#NUM!');

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

/** Arguments that each known function accepts. */
const ARITY: Record<string, readonly [number, number]> = {
  AVERAGE: [1, 255],
  COUNT: [1, 255],
  IF: [2, 3],
  MAX: [1, 255],
  MIN: [1, 255],
  SUM: [1, 255],
};

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
    const upper = name.toUpperCase();
    const arity = ARITY[upper];
    if (arity && (args.length < arity[0] || args.length > arity[1])) {
      return null;
    }
    return { type: 'call', name: arity ? upper : name, args };
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

/** The text of a formula, without its leading `=`, as Excel writes it. */
export function print(ast: Ast): string {
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
      return ast.operator + print(ast.operand);
    case 'percent':
      return `${print(ast.operand)}%`;
    case 'binary':
      return print(ast.left) + ast.operator + print(ast.right);
    case 'paren':
      return `(${print(ast.inner)})`;
    case 'call':
      return `${ast.name}(${ast.args.map(print).join(',')})`;
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

function areaOf(range: Extract<Ast, { type: 'range' }>): Area {
  const { start, end = range.start } = range;
  // Both ends of a range have rows, or neither has; columns alike.
  const [top, bottom] =
    start.row === null ? [0, SHEET_ROWS - 1] : order(start.row, end.row!);
  const [left, right] =
    start.col === null ? [0, SHEET_COLUMNS - 1] : order(start.col, end.col!);
  return { top, left, bottom, right };
}

const order = (a: number, b: number): [number, number] =>
  a <= b ? [a, b] : [b, a];

/** The letters of a column, counted from zero: `A`, ..., `Z`, `AA`, ... */
export function columnName(col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

// ---------------------------------------------------------------------------
// Evaluation

/** A reference as an operand evaluates to, before its cells are read. */
class Reference {
  constructor(
    readonly sheet: string | undefined,
    readonly area: Area,
  ) {}
}

type Result = Value | Reference;

/**
 * Calculates the value of a parsed formula. An empty result, as that of a
 * reference to an empty cell, is 0, as Excel shows it.
 */
export function evaluate(ast: Ast, context: Context): Value {
  return scalar(run(ast, context), context) ?? 0;
}

function run(ast: Ast, context: Context): Result {
  switch (ast.type) {
    case 'number':
    case 'string':
    case 'boolean':
    case 'error':
      return ast.value;
    case 'missing':
      return 0;
    case 'name':
      return error('#NAME?');
    case 'range':
      return new Reference(ast.sheet, areaOf(ast));
    case 'paren':
      return run(ast.inner, context);
    case 'unary': {
      const operand = scalar(run(ast.operand, context), context);
      if (ast.operator === '+') return operand;
      const number = toNumber(operand);
      return isError(number) ? number : -number;
    }
    case 'percent': {
      const number = toNumber(scalar(run(ast.operand, context), context));
      return isError(number) ? number : number / 100;
    }
    case 'binary':
      return ast.operator === ' '
        ? intersect(run(ast.left, context), run(ast.right, context))
        : binary(
            ast.operator,
            scalar(run(ast.left, context), context),
            scalar(run(ast.right, context), context),
          );
    case 'call': {
      const implementation = LIBRARY[ast.name];
      return implementation
        ? implementation(ast.args, context)
        : error('#NAME?');
    }
  }
}

/**
 * The one value of a result. A range where one value is expected gives
 * its cell in the formula's row or column, as Excel does.
 */
function scalar(result: Result, context: Context): Value {
  if (!(result instanceof Reference)) return result;
  const { sheet, area } = result;
  if (area.top === area.bottom && area.left === area.right) {
    return context.cell(sheet, area.top, area.left);
  }
  if (area.left === area.right && within(context.row, area.top, area.bottom)) {
    return context.cell(sheet, context.row, area.left);
  }
  if (area.top === area.bottom && within(context.col, area.left, area.right)) {
    return context.cell(sheet, area.top, context.col);
  }
  return VALUE;
}

const within = (n: number, low: number, high: number) => n >= low && n <= high;

function intersect(left: Result, right: Result): Result {
  if (!(left instanceof Reference)) return isError(left) ? left : VALUE;
  if (!(right instanceof Reference)) return isError(right) ? right : VALUE;
  if (left.sheet?.toLowerCase() !== right.sheet?.toLowerCase()) return VALUE;
  const a = left.area;
  const b = right.area;
  const area = {
    top: Math.max(a.top, b.top),
    left: Math.max(a.left, b.left),
    bottom: Math.min(a.bottom, b.bottom),
    right: Math.min(a.right, b.right),
  };
  if (area.top > area.bottom || area.left > area.right) {
    return error('#NULL!');
  }
  return new Reference(left.sheet, area);
}

function binary(operator: BinaryOperator, left: Value, right: Value): Value {
  if (isError(left)) return left;
  if (isError(right)) return right;
  if (operator === '&') return toText(left) + toText(right);
  if (
    operator === '=' ||
    operator === '<>' ||
    operator === '<' ||
    operator === '>' ||
    operator === '<=' ||
    operator === '>='
  ) {
    const order = compare(left, right);
    switch (operator) {
      case '=':
        return order === 0;
      case '<>':
        return order !== 0;
      case '<':
        return order < 0;
      case '>':
        return order > 0;
      case '<=':
        return order <= 0;
      default:
        return order >= 0;
    }
  }
  const a = toNumber(left);
  if (isError(a)) return a;
  const b = toNumber(right);
  if (isError(b)) return b;
  let result: number;
  switch (operator) {
    case '+':
      result = a + b;
      break;
    case '-':
      result = a - b;
      break;
    case '*':
      result = a * b;
      break;
    case '/':
      if (b === 0) return DIV0;
      result = a / b;
      break;
    default:
      if (a === 0 && b === 0) return NUM;
      if (a === 0 && b < 0) return DIV0;
      result = a ** b;
  }
  return Number.isFinite(result) ? result : NUM;
}

/**
 * The order of two values as Excel compares them: numbers before text
 * before booleans, text without regard to case, and an empty cell as the
 * empty value of the other's type.
 */
function compare(left: Value, right: Value): number {
  const a = left ?? emptyLike(right);
  const b = right ?? emptyLike(left);
  const rank = (value: Value) =>
    typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  // Numbers are equal when they agree to the 15 digits Excel keeps.
  return digits15(Number(a)) - digits15(Number(b));
}

const digits15 = (n: number) => Number(n.toPrecision(15));

function emptyLike(value: Value): Value {
  if (typeof value === 'string') return '';
  if (typeof value === 'boolean') return false;
  return 0;
}

/**
 * The number an operand stands for: an empty cell is 0, a boolean 1 or 0,
 * and text a number only if it reads as one.
 */
function toNumber(value: Value): number | ErrorValue {
  if (value === null) return 0;
  if (typeof value === 'boolean') return Number(value);
  if (typeof value === 'string') {
    const read = parseInput(value.trim()).value;
    return typeof read === 'number' ? read : VALUE;
  }
  return value;
}

/** The text an operand of `&` stands for: numbers to 15 digits. */
function toText(value: Exclude<Value, ErrorValue>): string {
  if (value === null) return '';
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'number') {
    return String(digits15(value)).toUpperCase();
  }
  return value;
}

function toBoolean(value: Value): boolean | ErrorValue {
  if (isError(value)) return value;
  if (typeof value === 'string') {
    const upper = value.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') return upper === 'TRUE';
    return VALUE;
  }
  return Boolean(value);
}

// ---------------------------------------------------------------------------
// Functions

type Implementation = (args: readonly Ast[], context: Context) => Result;

/**
 * The numbers of the arguments of a function such as `SUM`: in ranges,
 * only the numbers; typed directly, anything that reads as a number. The
 * first error found is given instead.
 */
function numbers(
  args: readonly Ast[],
  context: Context,
): number[] | ErrorValue {
  const found: number[] = [];
  for (const arg of args) {
    const result = run(arg, context);
    if (result instanceof Reference) {
      const values = context.cells(result.sheet, result.area);
      if (!Array.isArray(values)) return values as ErrorValue;
      for (const value of values) {
        if (typeof value === 'number') found.push(value);
        else if (isError(value)) return value;
      }
    } else {
      const number = toNumber(result);
      if (isError(number)) return number;
      found.push(number);
    }
  }
  return found;
}

function aggregate(
  combine: (values: number[]) => number | ErrorValue,
): Implementation {
  return (args, context) => {
    const values = numbers(args, context);
    if (!Array.isArray(values)) return values;
    const result = combine(values);
    return typeof result === 'number' && !Number.isFinite(result)
      ? NUM
      : result;
  };
}

const sum = (values: number[]) => values.reduce((total, n) => total + n, 0);

/** The extreme of numbers, without spreading a whole column as arguments. */
const extreme =
  (pick: (a: number, b: number) => number) => (values: number[]) =>
    values.length === 0 ? 0 : values.reduce((a, b) => pick(a, b));

const LIBRARY: Record<string, Implementation> = {
  SUM: aggregate(sum),
  AVERAGE: aggregate((values) =>
    values.length === 0 ? DIV0 : sum(values) / values.length,
  ),
  MIN: aggregate(extreme(Math.min)),
  MAX: aggregate(extreme(Math.max)),
  COUNT(args, context) {
    let count = 0;
    for (const arg of args) {
      const result = run(arg, context);
      if (result instanceof Reference) {
        const values = context.cells(result.sheet, result.area);
        if (!Array.isArray(values)) continue;
        for (const value of values) {
          if (typeof value === 'number') count++;
        }
      } else if (!isError(result) && !isError(toNumber(result))) {
        count++;
      }
    }
    return count;
  },
  IF(args, context) {
    const condition = toBoolean(scalar(run(args[0]!, context), context));
    if (typeof condition !== 'boolean') return condition;
    const branch = args[condition ? 1 : 2];
    return branch ? run(branch, context) : false;
  },
};
