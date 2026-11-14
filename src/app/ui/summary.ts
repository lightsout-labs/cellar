import { formatValue } from '../../engine/format';
import { cellName, quoteSheetName, type Workbook } from '../../engine/workbook';
import type { Selection } from '../grid/selection';

/** The figures the status bar shows for a selection, as text. */
export interface Summary {
  /** The average and sum of the numbers, or `null` when there are none. */
  readonly average: string | null;
  readonly count: string;
  readonly sum: string | null;
}

/**
 * The average, count and sum of the selected cells, as Excel's status bar
 * shows them, or `null` for a single cell or a selection of empty cells.
 */
export function summarize(
  workbook: Workbook,
  sheet: string,
  selection: Selection,
): Summary | null {
  const { ranges } = selection;
  const [first] = ranges;
  if (
    ranges.length === 1 &&
    first!.top === first!.bottom &&
    first!.left === first!.right
  ) {
    return null;
  }
  const { count, numbers, sum } = workbook.aggregate(sheet, ranges);
  if (count === 0) return null;
  return {
    average: numbers > 0 ? formatValue(sum / numbers) : null,
    count: String(count),
    sum: numbers > 0 ? formatValue(sum) : null,
  };
}

/**
 * The name of a cell in a circular reference, as the status bar shows it,
 * with its sheet when it is not the one shown; `null` when there is none.
 */
export function circularCell(workbook: Workbook, shown: string): string | null {
  const found = workbook.circularReference();
  if (!found) return null;
  const name = cellName(found.row, found.col);
  if (found.sheet === shown) return name;
  const sheet = workbook.sheets.find((each) => each.id === found.sheet)!;
  return `${quoteSheetName(sheet.name)}!${name}`;
}
