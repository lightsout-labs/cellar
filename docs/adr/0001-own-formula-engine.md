# 0001: Cellar has its own formula engine and workbook model

## Context

Formulas are the core of the product: parsing, at least 150 functions,
Excel's error values, dependency tracking, and rewriting references when
rows, columns and sheets move (R9, R11–R16). The engine must also run
without a DOM (R50), and the project is MIT licensed.

Options considered:

- **HyperFormula**. Complete and fast, but licensed under GPLv3 or a paid
  commercial license. Neither is compatible with distributing Cellar under
  MIT.
- **`fast-formula-parser` or a similar MIT parser**. It evaluates formulas
  but has no dependency graph, no reference rewriting and no undo, so most
  of the engine would still be ours, built around a parser we do not
  control and that is no longer maintained.
- **Our own engine.**

## Decision

Cellar implements its own parser, evaluator, function library and
dependency graph, in `src/engine/formula` and `src/engine/workbook`.

## Consequences

- No license conflict, and no third-party code in the hot path of
  recalculation.
- The parser's syntax tree serves evaluation, reference rewriting, fill,
  the colored reference highlights and autocomplete alike.
- The function library is the largest single piece of work in the project.
  It is spread over two slices and tested against the examples in Excel's
  documentation.
- Agreement with Excel is our responsibility, function by function.
- A later need for most of Excel's library (out of scope now) would mean
  writing it or revisiting this decision.
