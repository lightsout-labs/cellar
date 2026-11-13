# 0004: The engine runs on the main thread and recalculates on demand

## Context

When one cell changes, 100,000 dependent formulas must show new results
within 1 second while the interface stays responsive (R46). The grid reads
cells from the workbook on every frame (R45).

Options considered:

- **The engine in a web worker.** Recalculation can never block the
  interface, but every read becomes asynchronous: the grid needs a cache of
  the visible cells, edits need optimistic updates, and the workbook's
  state exists in two places that must agree.
- **The engine on the main thread, recalculating everything eagerly.**
  Simple, but a long chain of formulas blocks input until it finishes.
- **The engine on the main thread, recalculating on demand.** A change
  marks its dependents dirty, which is fast. A dirty cell is calculated
  when something reads it. The remaining dirty cells are calculated in
  short batches between frames.

## Decision

The workbook lives on the main thread. `execute` marks dependents dirty and
returns. `cell()` calculates a dirty cell, and the dirty cells it depends
on, before answering. The workbook finishes the rest in batches of a few
milliseconds through a scheduling function that the host supplies.

On-demand calculation follows dependencies with an explicit stack, never
with recursion. A sheet can hold a chain of 100,000 formulas in which each
depends on the one before, and reading the last one must calculate them
all; recursive evaluation would overflow the call stack.

## Consequences

- The cells on screen are correct in the frame after an edit, however many
  other cells depend on it, and reads stay synchronous everywhere.
- The engine holds the only copy of the workbook; there is no protocol and
  no cache to keep in step.
- The depth of a dependency chain is limited by memory, not by the call
  stack. The evaluator is harder to write than a recursive one, and the
  same explicit stack detects circular references.
- Reading the end of a long dirty chain calculates the whole chain in one
  step, so that read blocks for as long as the chain takes.
- One formula that is slow by itself, such as a lookup over a very large
  range, still blocks for its own duration.
- Opening a large file parses it on the main thread, so the interface does
  not respond during the seconds R47 allows.
- If either limit becomes a problem, moving the engine into a worker is
  possible, because it has no DOM dependency, but the grid would then need
  the asynchronous read path this decision avoids.
