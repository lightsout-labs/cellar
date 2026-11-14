import { SHEET_COLUMNS, type Value } from '../../engine/workbook';
import { COLUMN_WIDTH } from './geometry';

/** The space between the text of a cell and its borders, at 100%. */
export const TEXT_PADDING = 4;

/** The text of a cell as a row of the grid draws it. */
export interface RowText {
  readonly col: number;
  readonly text: string;
  readonly align: 'left' | 'right' | 'center';
  /** The first and last columns the text may cover. */
  readonly from: number;
  readonly to: number;
}

/**
 * The texts that the columns `left` to `right` of a row show, as in Excel:
 * numbers align right, text left, and booleans and errors in the middle,
 * and text too long for its cell runs over the empty cells to its right.
 * It includes text from a cell to the left of `left` that runs into them.
 * `measure` gives the width of a text at 100%.
 */
export function rowTexts(
  cell: (col: number) => { value: Value; display: string },
  left: number,
  right: number,
  measure: (text: string) => number,
): RowText[] {
  const texts: RowText[] = [];
  const add = (col: number) => {
    const { value, display } = cell(col);
    if (display === '') return;
    const align =
      typeof value === 'number'
        ? 'right'
        : typeof value === 'string'
          ? 'left'
          : 'center';
    let to = col;
    if (align === 'left') {
      const width = measure(display) + 2 * TEXT_PADDING;
      while (
        (to - col + 1) * COLUMN_WIDTH < width &&
        to + 1 < SHEET_COLUMNS &&
        cell(to + 1).display === ''
      ) {
        to++;
      }
    }
    if (to >= left) texts.push({ col, text: display, align, from: col, to });
  };
  // The nearest text to the left may run into the columns.
  let before = left - 1;
  while (before >= 0 && cell(before).display === '') before--;
  if (before >= 0) add(before);
  for (let col = left; col <= right; col++) add(col);
  return texts;
}
