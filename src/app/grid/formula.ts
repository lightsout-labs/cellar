import {
  argumentHint,
  functions,
  print,
  references as areasOf,
  tokenize,
  type Area,
  type Corner,
  type Token,
} from '../../engine/formula';

/** The number of colors references take in turn. */
export const REFERENCE_COLORS = 6;

/** A reference in the formula being typed, and where it is in the text. */
export interface Highlight {
  readonly start: number;
  readonly end: number;
  /** The same for the same reference, from 0 up to `REFERENCE_COLORS`. */
  readonly color: number;
  /** The sheet named, or none for the formula's own sheet. */
  readonly sheet?: string;
  readonly area: Area;
}

/** What helps while a formula is typed, with the caret at some place. */
export interface FormulaHelp {
  readonly references: readonly Highlight[];
  /** The functions whose names start with the name typed at the caret. */
  readonly completion: {
    readonly start: number;
    readonly end: number;
    readonly names: readonly string[];
  } | null;
  /** The arguments of the function call that holds the caret. */
  readonly hint: {
    readonly name: string;
    readonly arguments: readonly string[];
    /** Which argument the caret is in, or -1 past the last one. */
    readonly current: number;
  } | null;
}

/** A new text of the entry and where the caret goes. */
export interface Edit {
  readonly text: string;
  readonly caret: number;
}

const NOTHING: FormulaHelp = { references: [], completion: null, hint: null };

/** A name that may be the start of a function's, as `SU` or `LOG1`. */
const FUNCTION_START = /^[A-Za-z][A-Za-z0-9._]*$/;

/**
 * The help for an entry with the caret at `caret`: none unless it is a
 * formula, which starts with `=`.
 */
export function formulaHelp(text: string, caret: number): FormulaHelp {
  const tokens = formulaTokens(text);
  if (!tokens) return NOTHING;
  return {
    references: highlights(tokens),
    completion: completion(text, tokens, caret),
    hint: hint(tokens, caret),
  };
}

/**
 * Whether a reference picked on the grid can go at the caret: where an
 * operand starts, or in place of a reference that ends there.
 */
export function pickable(text: string, caret: number): boolean {
  const tokens = formulaTokens(text);
  if (!tokens) return false;
  const before = tokenBefore(tokens, caret);
  if (!before) return caret > 0;
  if (before.end < caret && text.slice(before.end, caret).trim() !== '') {
    return false; // The caret is inside a token.
  }
  if (before.kind === 'reference') return before.end === caret;
  return (
    before.kind === 'open' ||
    before.kind === 'comma' ||
    before.kind === 'operator'
  );
}

/**
 * Puts a reference at the caret, in place of a reference that ends there,
 * as a drag on the grid replaces the one it started.
 */
export function insertReference(
  text: string,
  caret: number,
  reference: string,
): Edit {
  const before = tokenBefore(formulaTokens(text) ?? [], caret);
  const start =
    before?.kind === 'reference' && before.end === caret ? before.start : caret;
  return replace(text, start, caret, reference);
}

/**
 * Cycles the reference at the caret through `$A$1`, `A$1`, `$A1` and `A1`,
 * as F4 does; `null` when the caret is not on a reference.
 */
export function cycleReference(text: string, caret: number): Edit | null {
  const token = formulaTokens(text)?.find(
    (each) =>
      each.kind === 'reference' && each.start <= caret && caret <= each.end,
  );
  if (token?.kind !== 'reference') return null;
  const { reference } = token;
  const [colAbsolute, rowAbsolute] = next(reference.start);
  const absolute = (corner: Corner): Corner => ({
    ...corner,
    rowAbsolute,
    colAbsolute,
  });
  const cycled = print({
    ...reference,
    start: absolute(reference.start),
    ...(reference.end && { end: absolute(reference.end) }),
  });
  return replace(text, token.start, token.end, cycled);
}

/** Puts a function and its opening parenthesis in place of the name typed. */
export function complete(text: string, caret: number, name: string): Edit {
  const tokens = formulaTokens(text);
  const found = tokens && completion(text, tokens, caret);
  const start = found ? found.start : caret;
  return replace(text, start, caret, `${name}(`);
}

/** The tokens of a formula, placed in the whole text, or `null`. */
function formulaTokens(text: string): Token[] | null {
  if (!text.startsWith('=')) return null;
  return tokenize(text.slice(1)).map((token) => ({
    ...token,
    start: token.start + 1,
    end: token.end + 1,
  }));
}

/** The last token that starts before the caret. */
function tokenBefore(tokens: Token[], caret: number): Token | undefined {
  return tokens.findLast((token) => token.start < caret);
}

function highlights(tokens: Token[]): Highlight[] {
  const colors = new Map<string, number>();
  const found: Highlight[] = [];
  for (const token of tokens) {
    if (token.kind !== 'reference') continue;
    const key = print(token.reference).toUpperCase();
    if (!colors.has(key)) colors.set(key, colors.size % REFERENCE_COLORS);
    const [{ sheet, area }] = areasOf(token.reference) as [
      { sheet?: string; area: Area },
    ];
    found.push({
      start: token.start,
      end: token.end,
      color: colors.get(key)!,
      ...(sheet !== undefined && { sheet }),
      area,
    });
  }
  return found;
}

function completion(
  text: string,
  tokens: Token[],
  caret: number,
): FormulaHelp['completion'] {
  const token = tokenBefore(tokens, caret);
  if (token?.end !== caret) return null;
  if (token.kind !== 'name' && token.kind !== 'reference') return null;
  const typed = text.slice(token.start, caret);
  // After a `$`, it is the column of a reference.
  if (!FUNCTION_START.test(typed) || text[token.start - 1] === '$') return null;
  const prefix = typed.toUpperCase();
  const names = functions.filter((name) => name.startsWith(prefix));
  return names.length > 0 ? { start: token.start, end: caret, names } : null;
}

function hint(tokens: Token[], caret: number): FormulaHelp['hint'] {
  const calls: { name: string | null; index: number }[] = [];
  let previous: Token | undefined;
  for (const token of tokens) {
    if (token.end > caret) break;
    if (token.kind === 'open') {
      calls.push({
        name: previous?.kind === 'function' ? previous.name : null,
        index: 0,
      });
    } else if (token.kind === 'close') {
      calls.pop();
    } else if (token.kind === 'comma' && calls.length > 0) {
      calls[calls.length - 1]!.index++;
    }
    previous = token;
  }
  const call = calls.findLast((each) => each.name !== null);
  const found = call && argumentHint(call.name!, call.index);
  return found ? { name: call.name!.toUpperCase(), ...found } : null;
}

/** The next form of a reference, by its first corner: `[col, row]` absolute. */
function next({ rowAbsolute, colAbsolute }: Corner): [boolean, boolean] {
  if (!rowAbsolute && !colAbsolute) return [true, true];
  if (rowAbsolute && colAbsolute) return [false, true];
  if (rowAbsolute) return [true, false];
  return [false, false];
}

function replace(
  text: string,
  start: number,
  end: number,
  inserted: string,
): Edit {
  return {
    text: text.slice(0, start) + inserted + text.slice(end),
    caret: start + inserted.length,
  };
}
