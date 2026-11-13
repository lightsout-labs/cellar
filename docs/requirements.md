# Requirements

Status: Draft

## Goal

Cellar is a spreadsheet application that provides the main features of
Microsoft Excel in the browser. It is for people who need to create, open,
edit and save spreadsheets, including real `.xlsx` files, without installing
software or creating an account. It is written in TypeScript, runs entirely
client-side, and is deployed to GitHub Pages. Its architecture keeps open the
option of later packaging the same code as Windows, Linux and macOS desktop
applications.

## Requirements

### Grid and editing

1. The app shows a grid of cells addressed by column letters and row
   numbers (A1 notation). Each sheet supports at least 100,000 rows and
   100 columns.
2. The user can select a cell, a rectangular range, whole rows, whole
   columns, the whole sheet, and multiple disjoint ranges (Ctrl/Cmd+click),
   by mouse and by keyboard (arrow keys, Shift+arrows, Ctrl/Cmd+arrows,
   Ctrl/Cmd+A, Home/End, Page Up/Down).
3. The user can edit a cell by typing, by double-click, by F2, or in the
   formula bar. Enter commits and moves down, Tab commits and moves right,
   Escape cancels.
4. Typed input is recognized as a number, date, time, percentage, boolean,
   text or formula (leading `=`), the way Excel recognizes it.
5. Undo and redo (Ctrl/Cmd+Z, Ctrl/Cmd+Y) cover every change the user can
   make to a workbook, with a history of at least 100 steps.
6. Cut, copy and paste work within the app, preserving formulas (with
   relative references adjusted) and formatting. Paste special offers
   values only, formulas only, and formatting only.
7. Copy and paste work to and from other applications through the system
   clipboard: a range copied from Excel or Google Sheets pastes as a grid of
   cells, and a range copied from Cellar pastes into them as a grid.
8. Dragging the fill handle fills a range by copying values, extending
   series (numbers, dates, weekday and month names), and adjusting relative
   references in formulas. Double-clicking the fill handle fills down to
   match the adjacent column.
9. The user can insert, delete, hide, unhide and resize rows and columns.
   Double-clicking a header border auto-fits the size to its content.
   Formulas that refer to moved cells are updated.
10. A workbook holds multiple sheets. The user can add, rename, delete,
    duplicate, reorder and recolor sheet tabs.

### Formulas

11. Formulas support the Excel operators (arithmetic, comparison, `&`
    concatenation, `%`), parentheses, relative, absolute and mixed
    references (`A1`, `$A$1`, `A$1`), ranges, whole-column and whole-row
    references, references to other sheets (`Sheet2!A1`), and named ranges.
12. At least 150 Excel functions are implemented, covering math and
    trigonometry, statistical, logical, text, date and time, lookup and
    reference (including `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`),
    information, and conditional aggregates (`SUMIF(S)`, `COUNTIF(S)`,
    `AVERAGEIF(S)`). Each returns the same result as Excel for the same
    inputs.
13. Formulas produce the Excel error values `#DIV/0!`, `#N/A`, `#NAME?`,
    `#NULL!`, `#NUM!`, `#REF!` and `#VALUE!` in the situations where Excel
    does, and errors propagate through dependent formulas.
14. When a cell changes, every cell that depends on it, directly or
    indirectly, is recalculated automatically, and only those cells.
15. A circular reference is detected, reported to the user, and does not
    hang or crash the app.
16. While a formula is being typed, the app offers function name
    autocomplete with an argument hint, lets the user insert references by
    clicking or dragging on the grid, and highlights each referenced range
    in a distinct color. F4 cycles a reference through absolute and
    relative forms.

### Formatting

17. The user can set, per cell: font family, font size, bold, italic,
    underline, strikethrough, text color, fill color, horizontal and
    vertical alignment, text wrap, and borders (per side, with style and
    color).
18. The user can apply number formats: general, number with decimal places
    and thousands separator, currency, accounting, percentage, scientific,
    date, time, text, and custom Excel format codes.
19. The user can merge and unmerge cells.
20. Conditional formatting supports rules based on cell value comparisons,
    text contains, duplicate values, top/bottom N, color scales and data
    bars, and updates as values change.
21. The user can freeze rows and columns so they stay visible while
    scrolling.
22. A format painter copies the formatting of one range to another.

### Data tools

23. The user can sort a range by one or more columns, ascending or
    descending, with an option to treat the first row as headers.
24. An auto-filter on a range's header row lets the user filter rows by
    selected values, by text conditions and by number conditions. Filtered
    rows are hidden, not deleted.
25. Find and replace (Ctrl/Cmd+F, Ctrl/Cmd+H) searches the current sheet or
    the whole workbook, with match-case and match-entire-cell options, in
    values or in formulas.
26. Data validation restricts a cell to a dropdown list, a number range, a
    date range or a text length, and rejects invalid input with a message.
27. The user can define, edit and delete named ranges, use them in
    formulas, and jump to them from the name box.

### Comments and links

28. The user can add, edit and delete a note on a cell. A cell with a note
    shows a marker, and the note appears on hover.
29. The user can insert a hyperlink in a cell to a web address or to a cell
    in the workbook. Clicking it follows the link.

### Files and storage

30. The app opens and saves `.xlsx` files. Cell values, formulas, the
    formatting of requirements 17–21, column widths and row heights, merged
    cells, multiple sheets, named ranges, data validation, notes and
    hyperlinks survive a save in Cellar followed by an open in Excel, and
    the reverse.
31. Opening an `.xlsx` file that contains features Cellar does not support
    (for example charts or pivot tables) does not fail: the supported
    content loads and the user is told what was dropped.
32. The app imports and exports `.csv` files, with correct handling of
    quoted fields, embedded commas and newlines, and UTF-8 text.
33. Workbooks are saved automatically to browser storage as the user works.
    Closing and reopening the browser restores them with no data loss.
34. A start screen lists the workbooks stored in the browser, and lets the
    user create a new one, open, rename, duplicate or delete a stored one,
    or open a file from disk.
35. After its first load the app works without a network connection.

### Printing

36. The user can print the current sheet or the current selection through
    the browser's print dialog (and so save it as PDF), with a choice of
    page orientation, paper size, margins, fit-to-width scaling, and
    whether to print gridlines and headers. Long sheets break across pages
    without cutting through a row.

### User interface

37. The interface uses the layout of Excel for the web: a tabbed ribbon
    toolbar, a formula bar with a name box, the grid, sheet tabs, and a
    status bar. Every feature above is reachable from the ribbon or a
    context menu.
38. Right-clicking a cell, row header, column header or sheet tab opens a
    context menu with the actions that apply to it.
39. The status bar shows sum, average and count for the current selection,
    and a zoom control from 50% to 200%.
40. The Excel keyboard shortcuts for the features above work, using Ctrl on
    Windows and Linux and Cmd on macOS.
41. The interface is available in English and Simplified Chinese. It
    defaults to the browser language, and the user can switch; the choice
    is remembered.
42. The interface has light and dark themes. It follows the system setting
    by default, and the user can switch; the choice is remembered.
43. The interface has a modern visual style of its own rather than a copy
    of Excel's: flat surfaces, generous and consistent spacing, rounded
    controls, a single consistent icon set, a restrained color palette with
    one accent color, and short, subtle transitions on menus, dialogs and
    hover states. The level of polish is that of Linear or Notion.
44. Controls are visually consistent, aligned and labeled, show hover,
    focus and disabled states, and have tooltips with their shortcut. No
    control is a placeholder that does nothing.

### Performance

45. A sheet with 100,000 populated rows and 100 columns scrolls at 60
    frames per second on a current mid-range laptop.
46. Editing one cell that 100,000 formula cells depend on shows the
    recalculated results within 1 second, and the interface stays
    responsive while it recalculates.
47. A 10 MB `.xlsx` file opens within 5 seconds.

### Quality and delivery

48. The formula engine, the file import and export, and the editing
    operations have automated tests. End-to-end tests in a real browser
    cover the main user flows. All tests run in CI on every pull request.
49. Every merge to `main` deploys the app to GitHub Pages automatically,
    and the deployed app works when served from the repository's
    sub-path.
50. The spreadsheet engine (workbook model, formulas, file formats) has no
    dependency on the browser DOM and runs under Node.js, as its tests
    demonstrate.
51. Everything the app needs from its host platform (opening and saving
    files, persistent storage, clipboard, printing) goes through one
    platform interface with a single web implementation, so that a desktop
    shell can supply another implementation without changes to the engine
    or the UI.
52. The repository has a `README.md` that states what Cellar is, links to
    the deployed app, lists its features, and gives the commands to
    install, run, test and build it.
53. The repository has a `LICENSE` file with the MIT license.
54. The repository has a `CLAUDE.md` that tells a coding agent how the
    project is laid out, the commands to run, test, lint and build it, and
    the conventions to follow. It is kept accurate as the project changes.

## Out of scope

- Desktop builds for Windows, Linux and macOS. Only the architecture that
  permits them (requirements 50 and 51) is delivered.
- Charts and pivot tables.
- Dynamic array formulas that spill (`FILTER`, `SORT`, `UNIQUE`,
  `SEQUENCE`), and financial and engineering function families beyond the
  agreed 150.
- Macros, VBA and scripting.
- Real-time collaboration, sharing, user accounts, cloud sync, and any
  server-side component.
- Phones, tablets and touch input.
- File formats other than `.xlsx` and `.csv` (for example `.xls`, `.ods`).
- Images, shapes and other drawing objects in sheets.
- Sheet and workbook password protection.
- Installable PWA packaging.

## Constraints

- Written in TypeScript.
- Web first: a static, client-side application hosted on GitHub Pages, with
  no backend.
- Supports the current versions of Chrome, Edge, Firefox and Safari on
  desktop, with mouse and keyboard, at window widths of 1024 px and above.
- Nothing in the engine or UI may prevent later packaging as a Windows,
  Linux and macOS desktop application.
- The product name is Cellar.

## Decisions

Answers the human gave to clarifying questions, where they refine a
requirement above.

- **Feature areas**: formatting and data tools are included; charts and
  pivot tables are not.
- **Formula library**: about 150 common functions, without dynamic arrays.
- **File formats**: `.xlsx` and `.csv`.
- **Storage**: automatic save to browser storage with a list of stored
  workbooks, plus opening and saving files on disk; no account or cloud.
- **Desktop versions**: architecture only; no desktop builds in this
  version.
- **UI quality bar**: the layout of Excel for the web, with a tabbed
  ribbon, in a modern visual style with the polish of Linear or Notion.
- **Scale**: 100,000 rows by 100 columns must stay smooth.
- **Extras**: English and Simplified Chinese interface, dark mode, printing
  and PDF export, cell notes and hyperlinks are all included.
- **Devices**: desktop browsers only.
- **Name**: Cellar.
- **Repository documents**: a README, an MIT license and a `CLAUDE.md` are
  required.
