# For AI coding agents

The documentation is published for agents as plain text:

- [llms.txt](https://ifmelate.github.io/SharpCell.Grid/llms.txt): an index of the pages with one line each.
- [llms-full.txt](https://ifmelate.github.io/SharpCell.Grid/llms-full.txt): all guides and the API reference in one file.

Every page of the site also has its Markdown source next to it.

## Rules to paste into an agent's instructions

```text
SharpCell.Grid shows SharpCell worksheets in Avalonia (SharpCell.Grid.Avalonia) or WPF (SharpCell.Grid.Wpf): new SpreadsheetView { Worksheet = sheet }.
The view reads cells when it paints: after changing cells or recalculating, call view.InvalidateVisual().
Sizes, hidden rows and columns and frozen panes are taken when Worksheet is set: after changing them, call view.ReloadLayout().
Rows and columns are numbered from 1: new CellPosition(row, column).
What a cell shows is SharpCell's Cell.Text and Cell.Style; change the workbook, not the view.
The selection is view.Selection, a GridSelection; view.ActiveCell is its active cell, and setting it selects that cell alone.
Copying (Avalonia: await view.CopyAsync(); WPF: view.Copy()) puts text and HTML for Excel on the clipboard; GridClipboard gives the same without a view.
People edit cells in the view as in Excel; set view.IsReadOnly = true to only show them. Each change is recalculated and raises view.Edited (a GridEdit with the cells changed and Recalculation.ChangedCells).
To change cells from code as typing does, set SharpCell's cell.Input, or use view.Editor (GridEditor) so the change can be undone.
For another UI framework, implement IGridCanvas and call GridPainter.Paint, Hit, CellNearest and ScrollIntoView, with GridSelection for the input.
```
