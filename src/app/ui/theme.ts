export type Theme = 'light' | 'dark';

/** The theme to start in: the remembered choice, or else the system's. */
export function startingTheme(
  remembered: string | null,
  systemIsDark: boolean,
): Theme {
  if (remembered === 'light' || remembered === 'dark') return remembered;
  return systemIsDark ? 'dark' : 'light';
}
