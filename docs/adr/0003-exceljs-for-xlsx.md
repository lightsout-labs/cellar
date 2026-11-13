# 0003: ExcelJS reads and writes `.xlsx` files

## Context

Values, formulas, formatting, column widths, merged cells, sheets, named
ranges, data validation, notes and hyperlinks must survive a round trip
with Excel (R30). Files with unsupported content must still open (R31). A
10 MB file must open within 5 seconds (R47). The code must run in the
browser and in Node (R50) and suit an MIT project.

Options considered:

- **SheetJS, community edition** (Apache 2.0). Fast and robust at reading,
  but it does not write cell styles; that is in the paid edition.
- **ExcelJS** (MIT). Reads and writes styles, number formats, merges,
  frozen panes, defined names, data validation, conditional formatting,
  notes and hyperlinks. Larger and slower than SheetJS, and released
  infrequently.
- **Our own reader and writer** over a zip library and an XML parser. Full
  control of speed and fidelity, at the cost of implementing the file
  format for every feature in R30.

## Decision

`src/engine/io` uses ExcelJS, loaded on demand when a file is opened or
saved so that it does not slow the first load.

## Consequences

- Every feature listed in R30 has a counterpart in ExcelJS, so the work is
  a mapping to and from the workbook snapshot.
- ExcelJS is hidden behind `readXlsx` and `writeXlsx`. No other module
  sees its types, so replacing it does not reach beyond `engine/io`.
- Speed is the risk: R47 is tested in the first slice that opens `.xlsx`
  files. If ExcelJS cannot meet it, the reader alone is replaced by our own
  streaming reader, and the writer stays.
- Content ExcelJS does not model, such as charts and pivot tables, is
  dropped on open, which is the behavior R31 asks for; `readXlsx` reports
  what it found and dropped.
