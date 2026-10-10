# Roadmap

Where SharpCell.Grid is going, in the order the work is planned. Each step ships when its tests
pass and what it shows is checked against Microsoft Excel. Dates are not promised; the project
has one maintainer.

## Now: 0.1 — a faithful viewer

A workbook from Excel looks the same in the grid as in Excel at 100% zoom.

| Step | What | Needs from SharpCell |
|---|---|---|
| Done | Headers, gridlines, `Cell.Text`, General alignment, active cell, mouse, arrows, Enter/Tab, wheel, numbers fitted to their column, row headers widened for long numbers | `Cell.Text` (0.5.0) |
| Done | Cell styles: fonts (name, size, bold, italic, underline, strikethrough, colour), solid fills, borders, horizontal and vertical alignment, wrapped text, indent | `Cell.Style`, theme and indexed colours resolved (0.6.0) |
| Done | Sheet geometry from the file: column widths, row heights, hidden columns and rows | `Worksheet` column widths, row heights, hidden columns (0.6.0) |
| Done | Merged cells, painted and navigated as one cell | `Worksheet.MergedAreas` (0.6.0) |
| Done | Text running on over empty neighbours, by its alignment, also from cells out of view | — |
| 5 | Frozen panes and the sheet's gridline setting | `Worksheet` frozen panes, gridlines shown |
| 6 | Range selection with Shift and the mouse; copy as tab-separated text and HTML that Excel pastes | — |
| 7 | WPF control over the same painter | — |
| Release | 0.1.0 on NuGet: `SharpCell.Grid`, `SharpCell.Grid.Avalonia`, `SharpCell.Grid.Wpf` | SharpCell 0.6.0 released |

## Next: 0.2 — editing

- Typing into a cell and the formula bar as a separate control; F2, Escape, Enter, Tab.
- Input read as Excel reads it: numbers, percentages, dates and times in the workbook's culture,
  `=` for formulas, `'` for text.
- Recalculation after each edit, repainting only what changed (SharpCell's changed cells).
- Undo and redo, Delete, paste of values from Excel.

## Later: 0.3 — working like Excel

- Resizing columns and rows with the mouse, double-click to fit.
- Inserting and deleting rows and columns, with formulas following (needs reference shifting in SharpCell).
- The fill handle for series and copied formulas.
- Zoom.
- Saving changed styles and geometry (needs the .xlsx writer in SharpCell to write styles).

## Quality, all along

- **Oracle for geometry and display:** desktop Excel exports `Range.Text`, `Left`, `Top`,
  `Width` and `Height` and `Range.Interior.Color` for test workbooks through SharpCell's
  `tools/excel-oracle`; the grid's numbers are compared with them. Checked this way before 0.1:
  the General shortening in narrow columns, theme colours with tints, indent steps and how
  centred text runs on when one side is blocked.
- **Snapshots** of rendered frames for every visible feature.
- **Benchmarks:** scrolling a million-row sheet, painting a full screen of styled cells.
- **NativeAOT and trimming** proved by a published smoke test, as in SharpCell.
- **Documentation site** with tested code snippets and `llms.txt`, as in SharpCell.

## Not planned

Charts, pivot tables, conditional formatting, data validation lists, comments, images and
macros. A workbook with them shows its cells; the rest is not drawn.
