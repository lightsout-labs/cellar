/** Where a tab is drawn, from the left of the tab strip. */
export interface TabBounds {
  readonly left: number;
  readonly width: number;
}

/**
 * The place in the tab order of the tab at `from`, dragged to `x`: after
 * every other tab whose middle `x` has passed.
 */
export function dropIndex(
  tabs: readonly TabBounds[],
  from: number,
  x: number,
): number {
  return tabs.filter(
    ({ left, width }, index) => index !== from && left + width / 2 < x,
  ).length;
}

/**
 * The colors a tab can be given, by the key of their names, chosen to read
 * in both themes.
 */
export const TAB_COLORS = [
  ['color.red', '#e5484d'],
  ['color.orange', '#f76b15'],
  ['color.yellow', '#ffc53d'],
  ['color.green', '#30a46c'],
  ['color.teal', '#12a594'],
  ['color.blue', '#0090ff'],
  ['color.purple', '#8e4ec6'],
  ['color.gray', '#8b8d98'],
] as const;
