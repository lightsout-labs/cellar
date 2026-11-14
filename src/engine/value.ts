/** An error value such as `#VALUE!`, as the result of a formula. */
export type ErrorValue = { readonly error: string };

/** What a cell holds or a formula returns. `null` is an empty cell. */
export type Value = number | string | boolean | null | ErrorValue;

export const VALUE_ERROR: ErrorValue = { error: '#VALUE!' };

export function isError(value: Value): value is ErrorValue {
  return typeof value === 'object' && value !== null;
}
