# Backlog

Status: Draft

Slices for Cellar, built in this order. Each one is testable end to end.

## S01: Walking skeleton: a formula calculated, saved and deployed

Status: Todo

Acceptance criteria:
- Typing `2` in A1 and `=A1+1` in A2 of the canvas grid shows `3` in A2. (R1,
  R3, R14)
- After a reload, the same cells show the same contents, restored from browser
  storage. (R33)
- The engine compiles without the DOM library and its unit tests run in Node.
  (R48, R50)
- Storage is reached only through the `Platform` interface, and lint fails on
  direct use of a browser storage, clipboard, file or print API anywhere else.
  (R51)
- `make check` and `make e2e` pass locally and in CI, where `check` and `e2e`
  are required on `main`. (R48)
- Every merge to `main` deploys the app to
  `https://lightsout-labs.github.io/cellar/`, and a smoke test against that
  address passes. (R49)
- The repository has a `README.md` with the deployed link and the install,
  run, test and build commands, the MIT `LICENSE`, and a `CLAUDE.md`. (R52,
  R53, R54)

## S02: The application shell, in two languages and two themes

Status: Todo

Acceptance criteria:
- The window shows a tabbed ribbon, a formula bar with a name box, the grid,
  sheet tabs and a status bar, laid out as in Excel for the web, at window
  widths from 1024 px. (R37)
- The interface starts in the browser's language, English or Simplified
  Chinese; the user can switch, every visible string changes, and the choice
  survives a reload. (R41)
- The interface starts in the system's light or dark theme; the user can
  switch, the grid and all controls change, and the choice survives a reload.
  (R42)
- The style is flat, with consistent spacing, rounded controls, one icon set,
  one accent color, and short transitions on menus, dialogs and hover states.
  (R43)
- Every control has hover, focus and disabled states and a tooltip, and a
  control whose feature is not built yet is absent rather than inert. (R44)
- The name box shows the address of the active cell, and the formula bar shows
  its contents. (R37)

## S03: Selecting cells and moving around a large sheet

Status: Todo

Acceptance criteria:
- A sheet has at least 100,000 rows and 100 columns, all reachable by
  scrolling. (R1)
- The mouse selects a cell, a range by dragging, a row or column by its
  header, the whole sheet by the corner, and further ranges with
  Ctrl/Cmd+click. (R2)
- Arrow keys, Shift+arrows, Ctrl/Cmd+arrows, Ctrl/Cmd+A, Home, End, Page Up
  and Page Down move or extend the selection as in Excel, and the grid scrolls
  to keep the active cell visible. (R2, R40)
- The status bar shows the sum, average and count of the selected cells. (R39)
- The zoom control in the status bar scales the grid from 50% to 200%. (R39)
- Scrolling a sheet of 100,000 by 100 populated cells draws every frame within
  the frame-time budget in the Chromium performance test. (R45)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S04: Editing cells and recognizing what is typed

Status: Todo

Acceptance criteria:
- Typing, double-clicking, F2 and the formula bar each start editing; Enter
  commits and moves down, Tab commits and moves right, Escape cancels. (R3,
  R40)
- Typed numbers, dates, times, percentages and booleans are stored as values
  of that type and displayed in the matching format; anything else is text.
  (R4)
- Numbers align right, text aligns left, and text longer than its cell runs
  over empty neighbors, as in Excel. (R4)
- Delete clears the selected cells. (R3)
- Every edit is saved automatically and is present after closing and reopening
  the browser. (R33)

## S05: Undo and redo

Status: Todo

Acceptance criteria:
- Ctrl/Cmd+Z undoes and Ctrl/Cmd+Y redoes each change made so far, and the
  ribbon has a button for each. (R5, R40)
- At least 100 consecutive changes can be undone in order. (R5)
- Undo restores the selection to the cells that changed. (R5)
- A new change after an undo discards the redo history. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S06: Multiple sheets

Status: Todo

Acceptance criteria:
- The user can add, rename, delete and duplicate sheets from the sheet tabs.
  (R10)
- Dragging a tab reorders the sheets, and a tab can be given a color. (R10)
- Right-clicking a sheet tab opens a context menu with these actions. (R38)
- The last remaining sheet cannot be deleted. (R10)
- Sheets and their order survive a reload. (R33)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S07: Formulas: operators, references, errors and recalculation

Status: Todo

Acceptance criteria:
- Formulas support `+ - * / ^`, the comparison operators, `&`, `%`, unary
  minus and parentheses, with Excel's precedence. (R11)
- Formulas accept relative, absolute and mixed references, ranges,
  whole-column and whole-row references, and references to other sheets,
  including quoted sheet names. (R11)
- `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT` and `IF` work over values and
  ranges. (R12)
- The seven Excel error values appear where Excel produces them and propagate
  through dependent formulas. (R13)
- Changing a cell updates every formula that depends on it, directly or
  indirectly, and calculates no others. (R14)
- A circular reference is reported to the user, and the app keeps working.
  (R15)
- After a change to a cell that 100,000 formulas depend on, the visible
  results are correct within 1 second and the interface accepts input
  meanwhile. (R46)

## S08: Opening and saving XLSX files: values, formulas and sheets

Status: Todo

Acceptance criteria:
- The user can open an `.xlsx` file from disk and sees its sheets, values and
  formulas. (R30)
- The user can save the workbook as an `.xlsx` file, and opening that file
  again gives the same sheets, values and formulas. (R30)
- A file with charts, pivot tables or formulas Cellar does not know still
  opens, and a message lists what was dropped. (R31)
- A file that is not a valid `.xlsx` gives an error message and leaves the
  current workbook untouched. (R30)
- A 10 MB `.xlsx` file opens within 5 seconds. (R47)
- File access goes through the `Platform` interface. (R51)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S09: Functions: math, statistics, logic and conditional totals

Status: Todo

Acceptance criteria:
- The math and trigonometry functions return Excel's results for the examples
  in Excel's documentation. (R12)
- The statistical functions do the same. (R12)
- The logical functions do the same, including `IFS`, `SWITCH`, `IFERROR` and
  `IFNA`. (R12)
- `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF` and `AVERAGEIFS`
  accept Excel's criteria syntax, including comparison operators and
  wildcards. (R12)
- A function called with the wrong number or type of arguments gives the error
  Excel gives. (R13)

## S10: Functions: text, dates, lookup and information

Status: Todo

Acceptance criteria:
- The text functions return Excel's results for the examples in Excel's
  documentation. (R12)
- The date and time functions do the same, using Excel's serial numbers. (R12)
- `VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH` and the other lookup and
  reference functions do the same, in exact and approximate modes. (R12)
- The information functions do the same. (R12)
- At least 150 functions are implemented in total, and a test counts them.
  (R12)

## S11: Help while typing a formula

Status: Todo

Acceptance criteria:
- Typing a function name after `=` shows a list of matching functions;
  choosing one inserts it. (R16)
- Inside a function's parentheses, a hint shows its arguments with the current
  one emphasized. (R16)
- Clicking or dragging on the grid while editing a formula inserts that
  reference, on another sheet too. (R16)
- Each reference in the formula being edited has its own color, and its range
  is outlined on the grid in that color. (R16)
- F4 cycles the reference at the cursor through `$A$1`, `A$1`, `$A1` and `A1`.
  (R16, R40)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S12: Cut, copy and paste inside Cellar

Status: Todo

Acceptance criteria:
- Ctrl/Cmd+C, X and V, the ribbon and the cell context menu copy, cut and
  paste the selection, with a moving border around the copied range. (R6, R38,
  R40)
- Pasted formulas have their relative references adjusted, and absolute ones
  kept. (R6)
- Paste carries formatting with the values. (R6)
- Cut and paste moves the cells and updates formulas that referred to them.
  (R6)
- Paste special offers values only, formulas only and formatting only. (R6)
- Pasting one cell into a larger selection fills the selection. (R6)
- Each of these changes can be undone and redone. (R5)

## S13: Copy and paste with other applications

Status: Todo

Acceptance criteria:
- A range copied in Excel or Google Sheets pastes into Cellar as the same grid
  of cells, with typed values recognized. (R7)
- A range copied in Cellar is placed on the system clipboard as tab-separated
  text and as an HTML table, which Excel and Google Sheets paste as a grid.
  (R7)
- Plain multi-line text pastes one line per row. (R7)
- The clipboard is reached only through the `Platform` interface. (R51)
- Each of these changes can be undone and redone. (R5)

## S14: The fill handle

Status: Todo

Acceptance criteria:
- Dragging the fill handle of a selection copies its values and formats into
  the dragged range, in any of the four directions. (R8)
- Formulas filled this way have their relative references adjusted. (R8)
- A selection of numbers or dates that forms a series is extended as a series,
  and so are weekday and month names. (R8)
- Double-clicking the fill handle fills down as far as the adjacent column has
  data. (R8)
- Ctrl/Cmd+D and Ctrl/Cmd+R fill down and right. (R40)
- Each of these changes can be undone and redone. (R5)

## S15: Rows and columns

Status: Todo

Acceptance criteria:
- The user can insert and delete rows and columns, and formulas that refer to
  the moved cells still refer to them; a formula that referred to a deleted
  cell shows `#REF!`. (R9)
- Dragging a header border resizes the row or column, and double-clicking it
  fits the size to the content. (R9)
- Rows and columns can be hidden and shown again. (R9)
- Right-clicking a row or column header opens a context menu with these
  actions. (R38)
- These actions apply to all selected rows or columns at once. (R9)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S16: Fonts, colors and alignment

Status: Todo

Acceptance criteria:
- The user can set the font family, the font size, bold, italic, underline and
  strikethrough of the selected cells. (R17)
- The user can set the text color and the fill color from a palette or a
  custom color. (R17)
- The user can set horizontal and vertical alignment and turn text wrap on; a
  wrapped row grows to fit its text. (R17)
- Ctrl/Cmd+B, I and U toggle bold, italic and underline. (R40)
- The ribbon controls show the format of the active cell. (R37)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S17: Borders, merged cells and the format painter

Status: Todo

Acceptance criteria:
- The user can set a border on any side of the selected cells, or all, outside
  or none, with a line style and a color. (R17)
- The user can merge the selected cells into one and unmerge them; the merged
  cell keeps the top-left value. (R19)
- Selection and keyboard movement treat a merged cell as one cell. (R2, R19)
- The format painter copies all formatting of the selected range to the next
  range the user selects. (R22)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S18: Number formats

Status: Todo

Acceptance criteria:
- The ribbon applies the general, number, currency, accounting, percentage,
  scientific, date, time and text formats. (R18)
- The user can increase and decrease decimal places and toggle the thousands
  separator. (R18)
- A dialog accepts a custom Excel format code, including sections for negative
  numbers and zero, colors and literal text, and previews the result. (R18)
- A cell with the text format keeps what is typed as text, including leading
  zeros. (R18)
- A number too wide for its column shows `####`. (R18)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S19: Freeze panes

Status: Todo

Acceptance criteria:
- The user can freeze the top row, the first column, or all rows above and
  columns left of the active cell, and unfreeze them. (R21)
- Frozen rows and columns stay in place while the rest scrolls, with a line
  marking the boundary. (R21)
- Selection, editing and the fill handle work across the boundary. (R21)
- Frozen panes survive a reload. (R33)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S20: Conditional formatting: value rules

Status: Todo

Acceptance criteria:
- The user can add a rule to a range that formats cells greater than, less
  than, between or equal to a value, or whose text contains a string. (R20)
- The user can add a rule for duplicate values and for the top or bottom N
  values. (R20)
- A rule's format sets the text color, fill color and font style, and shows
  over the cell's own format. (R20)
- The formatting updates when a value changes, including the result of a
  formula. (R20)
- A panel lists the rules of the sheet and lets the user edit, reorder and
  delete them. (R20)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S21: Conditional formatting: color scales and data bars

Status: Todo

Acceptance criteria:
- The user can apply a two-color or three-color scale to a range, and each
  cell is shaded by where its value lies in the range. (R20)
- The user can apply data bars to a range, and each cell shows a bar whose
  length is in proportion to its value. (R20)
- Scales and bars update when values change. (R20)
- They appear in the rules panel and can be edited and deleted there. (R20)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S22: Sorting

Status: Todo

Acceptance criteria:
- The user can sort the selected range ascending or descending by the column
  of the active cell. (R23)
- A dialog sorts by several columns, each ascending or descending, with an
  option to treat the first row as headers. (R23)
- Numbers sort before text, text sorts without regard to case, and empty cells
  sort last, as in Excel. (R23)
- Whole rows of the range move together, with their formats, and formulas
  inside them are adjusted. (R23)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S23: Filtering

Status: Todo

Acceptance criteria:
- The user can turn on a filter for a range, which puts a dropdown button in
  each header cell. (R24)
- The dropdown filters rows by a checklist of the column's values, with a
  search box. (R24)
- The dropdown filters by text conditions (equals, contains, begins with, ends
  with) and number conditions (equals, greater than, less than, between).
  (R24)
- Rows that do not match are hidden, their row numbers show the gap, and
  clearing the filter shows them again unchanged. (R24)
- Filters on several columns combine, and a filtered column's button shows
  that it is filtered. (R24)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S24: Find and replace

Status: Todo

Acceptance criteria:
- Ctrl/Cmd+F opens a find panel that moves through the matching cells and
  shows how many there are. (R25, R40)
- Ctrl/Cmd+H adds replace and replace all. (R25, R40)
- The search covers the current sheet or the whole workbook. (R25)
- Options match case, match the entire cell, and search in values or in
  formulas. (R25)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S25: Data validation

Status: Todo

Acceptance criteria:
- The user can restrict the selected cells to a list, given as typed items or
  as a range; such a cell shows a dropdown to choose from. (R26)
- The user can restrict cells to a number range, a date range or a text
  length. (R26)
- Input that breaks the rule is rejected with a message. (R26)
- The user can view, change and remove the rule of a cell. (R26)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S26: Named ranges

Status: Todo

Acceptance criteria:
- The user can name the selected range by typing a name in the name box. (R27)
- A manager lists the names and lets the user edit and delete them. (R27)
- A name can be used in a formula in place of its range and is offered by
  formula autocomplete. (R11, R27)
- Choosing a name in the name box, or typing a cell address there, selects
  that range. (R27)
- A formula using a deleted name shows `#NAME?`, and names follow their cells
  when rows and columns are inserted or deleted. (R13, R27)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S27: Notes and hyperlinks

Status: Todo

Acceptance criteria:
- The user can add, edit and delete a note on a cell from the ribbon and the
  cell context menu. (R28, R38)
- A cell with a note shows a corner marker, and the note appears when the
  pointer rests on the cell. (R28)
- The user can insert a hyperlink to a web address or to a cell in the
  workbook, with display text, and edit or remove it. (R29)
- Clicking a web link opens it in a new tab, and clicking a cell link selects
  that cell. (R29)
- Each of these changes can be undone and redone. (R5)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S28: XLSX files: formatting and everything else

Status: Todo

Acceptance criteria:
- Fonts, colors, alignment, wrap, borders and number formats survive saving
  and reopening an `.xlsx` file. (R30)
- Column widths, row heights, hidden rows and columns, merged cells and frozen
  panes survive it. (R30)
- Conditional formats, data validation and named ranges survive it. (R30)
- Notes and hyperlinks survive it. (R30)
- The fixture files saved by Excel, covering each of these features, open with
  every feature shown correctly. (R30)
- A feature of the file that Cellar does not support is named in the message
  shown after opening. (R31)

## S29: CSV files

Status: Todo

Acceptance criteria:
- The user can open a `.csv` file as a new workbook, with typed values
  recognized. (R32)
- Quoted fields, commas and line breaks inside fields, doubled quotes and
  UTF-8 text, with or without a byte order mark, are read correctly. (R32)
- The user can export the current sheet as `.csv`, with fields quoted where
  needed and formulas exported as their results. (R32)
- Exporting and reopening a sheet gives the same values. (R32)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S30: The start screen and stored workbooks

Status: Todo

Acceptance criteria:
- Opening the app shows a start screen that lists the workbooks stored in the
  browser, most recent first, with their modification times. (R34)
- From it the user can create a new workbook, open a stored one, and open a
  file from disk. (R34)
- The user can rename, duplicate and delete a stored workbook, and deleting
  asks for confirmation. (R34)
- The workbook's name is shown in the title area and can be edited there; the
  page title follows it. (R34)
- Closing the browser in the middle of work and reopening the workbook loses
  nothing. (R33)
- A failure to save, such as storage being full, is shown to the user. (R33)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)

## S31: Working offline

Status: Todo

Acceptance criteria:
- After one visit, the app loads and opens stored workbooks with the network
  disconnected. (R35)
- Opening and saving `.xlsx` and `.csv` files works offline. (R35)
- After a new version is deployed, the next visit loads it, and tells the user
  if a reload is needed. (R35, R49)

## S32: Printing and saving as PDF

Status: Todo

Acceptance criteria:
- Ctrl/Cmd+P and the ribbon open a print panel for the current sheet or the
  current selection. (R36, R40)
- The user can choose orientation, paper size and margins, and the printed
  page follows them. (R36)
- The user can fit the sheet to the page width, and choose whether gridlines
  and row and column headers print. (R36)
- A sheet longer than a page breaks between rows, never through one. (R36)
- The printed cells show their formats, borders, merges and column widths.
  (R36)
- Printing goes through the `Platform` interface, and the browser's dialog can
  save the result as PDF. (R36, R51)
- Its controls are in the ribbon or a context menu, are translated in both
  languages, look right in both themes, and have hover, focus and disabled
  states and tooltips. (R37, R41, R42, R43, R44)
