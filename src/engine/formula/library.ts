import { CONDITIONAL } from './conditional';
import { LOGIC } from './logic';
import { MATH } from './math';
import { STATISTICS } from './statistics';
import type { Spec } from './values';

/** Every function Gridworks calculates, by its name in capitals. */
export const LIBRARY: Readonly<Record<string, Spec>> = {
  ...MATH,
  ...STATISTICS,
  ...LOGIC,
  ...CONDITIONAL,
};
