# For AI coding agents

The documentation is published for agents as plain text:

- [llms.txt](https://ifmelate.github.io/SharpCell.Grid/llms.txt): an index of the pages with one line each.
- [llms-full.txt](https://ifmelate.github.io/SharpCell.Grid/llms-full.txt): all guides and the API reference in one file.

Every page of the site also has its Markdown source next to it.

## Rules to paste into an agent's instructions

```text
SharpCell.Grid shows SharpCell worksheets in Avalonia: new SpreadsheetView { Worksheet = sheet }.
The view reads cells when it paints: after changing cells or recalculating, call view.InvalidateVisual().
Sizes, hidden rows and columns and frozen panes are taken when Worksheet is set: after changing them, call view.ReloadLayout().
Rows and columns are numbered from 1: new CellPosition(row, column).
What a cell shows is SharpCell's Cell.Text and Cell.Style; change the workbook, not the view.
The grid is for viewing; it does not edit cells yet.
For another UI framework, implement IGridCanvas and call GridPainter.Paint, HitTest and ScrollIntoView.
```
