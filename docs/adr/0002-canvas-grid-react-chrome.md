# 0002: The grid is drawn on a canvas, and React builds everything around it

## Context

A sheet of 100,000 by 100 populated cells must scroll at 60 frames per
second (R45). The interface around the grid is a ribbon, menus, dialogs and
panels with many stateful controls, in two languages and two themes, at the
polish of Linear or Notion (R37–R44).

Options considered for the grid:

- **DOM cells, virtualized.** Each scroll step creates, moves and styles
  hundreds of elements. Merged cells, frozen panes, borders and overflowing
  text make the layout work heavy, and frame times vary by browser.
- **An existing grid component** (AG Grid, Handsontable, Glide Data Grid).
  They are table components: Excel behaviors such as merged cells, the fill
  handle, reference picking while typing a formula, and per-side borders
  are missing or need their paid editions.
- **Our own canvas renderer.** Draws only the visible cells in one pass.

Options considered for the rest: React, a smaller framework (Preact, Solid,
Svelte), or no framework.

## Decision

The grid is one canvas element drawn by `src/app/grid`, which also handles
its mouse and keyboard input and places a single DOM text editor over the
cell being edited. Everything else is React, with Radix UI primitives for
menus, dialogs, tooltips and tabs.

## Consequences

- Drawing cost depends on the size of the window, not of the sheet.
- We write hit testing, text measurement, clipping and high-density display
  handling ourselves.
- Canvas content is invisible to screen readers and to DOM-based test
  selectors. End-to-end tests drive the grid with the keyboard and the
  mouse and read results from the formula bar, the status bar and a test
  hook that reports a cell's displayed text.
- React is the most widely known choice, and the Radix primitives depend on
  it. The grid does not depend on React beyond a thin wrapper component, so
  the framework could be replaced without touching the renderer.
