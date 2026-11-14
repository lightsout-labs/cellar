import { describe, expect, it } from 'vitest';
import { startingTheme } from './theme';

describe('startingTheme', () => {
  it('follows the system setting', () => {
    expect(startingTheme(null, true)).toBe('dark');
    expect(startingTheme(null, false)).toBe('light');
  });

  it('prefers the remembered choice, and ignores one it does not know', () => {
    expect(startingTheme('light', true)).toBe('light');
    expect(startingTheme('dark', false)).toBe('dark');
    expect(startingTheme('sepia', true)).toBe('dark');
  });
});
