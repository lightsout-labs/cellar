import { describe, expect, it } from 'vitest';
import { dropIndex } from './sheets';

// Three tabs side by side, 100 pixels wide.
const tabs = [
  { left: 0, width: 100 },
  { left: 100, width: 100 },
  { left: 200, width: 100 },
];

describe('dropIndex', () => {
  it('places a dragged tab after every other tab whose middle it passed', () => {
    expect(dropIndex(tabs, 0, 140)).toBe(0);
    expect(dropIndex(tabs, 0, 160)).toBe(1);
    expect(dropIndex(tabs, 0, 260)).toBe(2);
    expect(dropIndex(tabs, 2, 160)).toBe(2);
    expect(dropIndex(tabs, 2, 140)).toBe(1);
    expect(dropIndex(tabs, 2, 40)).toBe(0);
    expect(dropIndex(tabs, 1, 500)).toBe(2);
    expect(dropIndex(tabs, 1, -50)).toBe(0);
  });
});
