# Roadmap

Where SharpCell.Grid is going, in the order the work is planned. Each step ships when its tests
pass and what it shows is checked against Microsoft Excel. Dates are not promised; the project
has one maintainer.

## Released: 0.1 — a faithful viewer

A workbook from Excel looks the same in the grid as in Excel at 100% zoom.

| Step | What | Needs from SharpCell |
|---|---|---|
| Done | Headers, gridlines, `Cell.Text`, General alignment, active cell, mouse, arrows, Enter/Tab, wheel, numbers fitted to their column, row headers widened for long numbers | `Cell.Text` (0.5.0) |
| Done | Cell styles: fonts (name, size, bold, italic, underline, strikethrough, colour), solid fills, borders, horizontal and vertical alignment, wrapped text, indent | `Cell.Style`, theme and indexed colours resolved (0.6.0) |
| Done | Sheet geometry from the file: column widths, row heights, hidden columns and rows | `Worksheet` column widths, row heights, hidden columns (0.6.0) |
| Done | Merged cells, painted and navigated as one cell | `Worksheet.MergedAreas` (0.6.0) |
| Done | Text running on over empty neighbours, by its alignment, also from cells out of view | — |
| Done | Frozen panes and the sheet's gridline setting | `Worksheet.FrozenRows`, `FrozenColumns`, `ShowGridlines` (0.6.0) |
| Done | A live demo in the browser through WebAssembly, on the documentation site | — |
| Released | 0.1.0 on NuGet: `SharpCell.Grid`, `SharpCell.Grid.Avalonia` | SharpCell 0.6.0 |

## Now: 0.2 — selection and WPF

| Step | What | Needs from SharpCell |
|---|---|---|
| Done | Range selection: the mouse with dragging and Shift, Shift with the arrows, whole columns and rows from their headers and with Ctrl+Space and Shift+Space, the whole sheet; Enter and Tab inside a selection | — |
| Done | Copy as tab-separated text and as HTML with styles, values as shown | — |
| Done | A WPF control over the same painter: `SharpCell.Grid.Wpf`, tested on Windows in CI | — |

## Now: 0.3 — editing

| Step | What | Needs from SharpCell |
|---|---|---|
| Done | Typing into a cell, F2 and double-click, Enter, Tab, Escape, arrows, Ctrl+Enter, Alt+Enter; `IsReadOnly` | — |
| Done | Input read as Excel reads it: numbers, percentages, currency, dates and times in the workbook's culture, `=` for formulas, `'` for text, the number format a General cell takes | `Cell.Input` (0.7.0) |
| Done | Recalculation after each edit, with the cells it changed (`Edited`) | `RecalculationResult.ChangedCells` (0.5.0) |
| Done | Undo and redo, Delete, paste of text copied from Excel | — |
| Done | The formula bar as a control of its own, with the name box | — |
| Done | Ctrl with the arrows, Home, Ctrl+Home, Ctrl+End, Page Up and Page Down | — |
| Later | Ctrl+A on the cells around the active one first; references picked with the mouse while typing a formula | — |

## Later: 0.4 — working like Excel

- Resizing columns and rows with the mouse, double-click to fit.
- Inserting and deleting rows and columns, with formulas following (needs reference shifting in SharpCell).
- The fill handle for series and copied formulas.
- Zoom.
- Saving changed styles and geometry (needs the .xlsx writer in SharpCell to write styles).

## Quality, all along

- **Oracle for geometry and display:** desktop Excel exports `Range.Text`, `Left`, `Top`,
  `Width` and `Height` and `Range.Interior.Color` for test workbooks through SharpCell's
  `tools/excel-oracle`; the grid's numbers are compared with them. To be checked this way:
  the General shortening in narrow columns, theme colours with tints, indent steps and how
  centred text runs on when one side is blocked.
- **Snapshots** of rendered frames for every visible feature.
- **Benchmarks:** scrolling a million-row sheet, painting a full screen of styled cells.
- **NativeAOT and trimming** proved by a published smoke test, as in SharpCell.
- **Documentation site** with tested code snippets and `llms.txt`, as in SharpCell.

## Not planned

Charts, pivot tables, conditional formatting, data validation lists, comments, images and
macros. A workbook with them shows its cells; the rest is not drawn.
