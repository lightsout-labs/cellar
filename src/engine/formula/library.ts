import { CONDITIONAL } from './conditional';
import { DATE_FUNCTIONS } from './date';
import { INFORMATION } from './information';
import { LOGIC } from './logic';
import { LOOKUP_FUNCTIONS } from './lookup';
import { MATH } from './math';
import { STATISTICS } from './statistics';
import { TEXT_FUNCTIONS } from './text';
import type { Spec } from './values';

/** Every function Gridworks calculates, by its name in capitals. */
export const LIBRARY: Readonly<Record<string, Spec>> = {
  ...MATH,
  ...STATISTICS,
  ...LOGIC,
  ...CONDITIONAL,
  ...TEXT_FUNCTIONS,
  ...DATE_FUNCTIONS,
  ...LOOKUP_FUNCTIONS,
  ...INFORMATION,
};
