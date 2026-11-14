/** An error value such as `#VALUE!`, as the result of a formula. */
export type ErrorValue = { readonly error: string };

/** What a cell holds or a formula returns. `null` is an empty cell. */
export type Value = number | string | boolean | null | ErrorValue;

export const VALUE_ERROR: ErrorValue = { error: '#VALUE!' };

/** Whether a value, or anything a function gives instead, is an error. */
export function isError(value: unknown): value is ErrorValue {
  return typeof value === 'object' && value !== null && 'error' in value;
}
