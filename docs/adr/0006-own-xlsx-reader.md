# 0006: Gridworks reads `.xlsx` files with its own reader, and writes them with ExcelJS

## Context

ADR 0003 chose ExcelJS for reading and writing `.xlsx` files, and named
the risk: a 10 MB file must open within 5 seconds (R47). Measured in
Node.js, ExcelJS took 4.5 seconds to load an 11 MB file of 150,000 rows of
ten numbers and texts and a formula, before any mapping to a snapshot; a
browser is slower still. The same file unzipped with JSZip and scanned
with regular expressions took 0.8 seconds.

Options considered:

- **Keep ExcelJS for both**, and accept that a large file may take more
  than 5 seconds. That fails R47.
- **SheetJS for reading**, ExcelJS for writing. Faster than ExcelJS, but a
  second large library beside the one kept for writing, with its own
  model of a workbook to map from.
- **Our own reader** over JSZip, scanning the parts it needs, and
  ExcelJS for writing, as ADR 0003 foresaw.

## Decision

`readXlsx` is our own reader over JSZip. It reads the workbook,
relationship, shared string, style and worksheet parts with regular
expressions over the XML text, and skips every other part. `writeXlsx`
keeps ExcelJS. ADR 0003 stands for writing.

## Consequences

- A 10 MB file opens in about 2 seconds in Node.js, reading and building
  the workbook together, and the unit test holds it under 5.
- Each feature that R30 adds to the files (fonts, fills, borders, merges,
  widths, validation, notes, links) needs reading code of our own in S28,
  where ExcelJS would have given it; writing still comes from ExcelJS.
- The reader assumes XML as Excel and the common writers produce it:
  attributes in quotes, no CDATA in cells. A file it cannot read is
  reported as not a valid `.xlsx` file.
- JSZip is loaded with the reader, on demand, like ExcelJS.
