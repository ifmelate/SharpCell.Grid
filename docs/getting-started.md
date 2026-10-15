# Getting started

## Install

```bash
dotnet add package SharpCell.Grid.Avalonia   # the control; brings SharpCell.Grid and SharpCell
dotnet add package SharpCell.Xlsx            # to read .xlsx files
```

The packages target .NET 8 and .NET 10 and Avalonia 12. For WPF, see [WPF](#wpf) below.

## Show a workbook

```csharp snippet=show-workbook
Workbook workbook = XlsxReader.Load(path);
workbook.Recalculate();

var view = new SpreadsheetView { Worksheet = workbook["Budget"] };
```

`SpreadsheetView` is an Avalonia control: put it in a window or any panel, and it takes the space
it is given. `XlsxReader.Load` reads values, formulas, the results Excel saved, number formats,
styles and the sheet's sizes; `Recalculate` calculates every formula with SharpCell.

A runnable window is in [samples/SharpCell.Grid.Sample](../samples/SharpCell.Grid.Sample):

```bash
dotnet run --project samples/SharpCell.Grid.Sample -f net10.0 -- path/to/file.xlsx
```

## After changes

```csharp snippet=refresh
sheet["B2"].Value = 99;
workbook.Recalculate();
view.InvalidateVisual();     // values, formulas and styles show at the next paint

sheet.SetColumnWidth(2, 20);
sheet.FrozenRows = 1;
view.ReloadLayout();         // sizes and frozen panes are taken again
```

The view reads cells each time it paints, but it does not know when a workbook changes. Column
widths, row heights, hidden rows and columns and frozen panes are taken when `Worksheet` is set,
and again with `ReloadLayout`.

## WPF

```bash
dotnet add package SharpCell.Grid.Wpf   # the WPF control; brings SharpCell.Grid and SharpCell
dotnet add package SharpCell.Xlsx
```

`SharpCell.Grid.Wpf` has the same `SpreadsheetView`, as a WPF `FrameworkElement` in the namespace
`SharpCell.Grid.Wpf`, for `net8.0-windows` and `net10.0-windows`. In XAML:

```xml
<Window xmlns:grid="clr-namespace:SharpCell.Grid.Wpf;assembly=SharpCell.Grid.Wpf" ...>
    <grid:SpreadsheetView x:Name="View" />
</Window>
```

```csharp snippet=show-workbook-wpf
Workbook workbook = XlsxReader.Load(path);
workbook.Recalculate();
View.Worksheet = workbook.Sheets[0];   // View: <grid:SpreadsheetView x:Name="View" /> in MainWindow.xaml
```

It shows, selects and copies as the Avalonia control does, and has the same properties: `Worksheet`,
`Selection`, `ActiveCell`, `Offset` (dependency properties), `Painter`, `ReloadLayout()`; copying
is `Copy()`, also bound to `ApplicationCommands.Copy`, and Ctrl+A to `ApplicationCommands.SelectAll`.
A runnable window is in [samples/SharpCell.Grid.WpfSample](../samples/SharpCell.Grid.WpfSample).

## Editing

```csharp snippet=editing
var bar = new FormulaBar { View = view };   // Excel's formula bar, as a control of its own
view.Edited += (_, edit) =>                 // after each change, the workbook recalculated
    changed = edit.Recalculation.ChangedCells;
view.IsReadOnly = true;                      // to only show and select
```

Cells are edited as in Excel: typing into a cell, F2 or a double-click, Enter, Tab and Escape,
Delete, paste with Ctrl+V, undo and redo. What is typed is read as Excel reads it, in the
workbook's culture, and the workbook is recalculated after each change. `view.Editor` is the
`GridEditor` behind it, with the history to undo; `view.Edited` tells about each change.
`FormulaBar` shows and edits the active cell's formula above the grid. See
[What the grid shows](guides/showing-workbooks.md#editing) for every key.

## Sheets and formatting

```csharp snippet=sheets-and-formatting
var tabs = new SheetTabs { View = view };   // Excel's sheet tabs, under the view
view.ToggleBold();                          // as Ctrl+B
view.SetFill(CellColor.FromRgb(0xFFF2CC));  // as Fill Color
view.SetNumberFormat("0.0%");               // as the Number Format box
view.ShowNextSheet();                       // as Ctrl+Page Down: the next sheet
```

`SheetTabs` is Excel's row of sheet tabs as a control of its own: clicking a tab shows its sheet,
and + adds one; Ctrl+Page Down and Ctrl+Page Up go through the sheets too, and each sheet keeps its
selection. The view's formatting methods (`ToggleBold`, `SetFill`, `SetBorders`, `SetNumberFormat`
and the others) do what Excel's Home tab does to the selected cells, for an application's toolbar,
and Excel's shortcuts such as Ctrl+B and Ctrl+Shift+% do the same. See
[What the grid shows](guides/showing-workbooks.md#formatting-cells) for every key.

## Rows, columns, filling and zoom

```csharp snippet=sheet-commands
view.Zoom = 1.5;          // 150%, as Ctrl or Cmd with the wheel
view.Editor!.Fill(sheet, new CellArea(1, 1, 2, 1), new CellArea(1, 1, 5, 1));   // as the fill handle: 1 to 5
view.InsertRows();        // a row before the selected one, A2, as Ctrl+Plus
view.FillDown();          // the row above copied into it, as Ctrl+D
```

Columns and rows are resized by dragging the edges of their headers and fitted to their content
by double-clicking them. The fill handle on the selection's corner continues series and copies
formulas, right-clicking a header gives Excel's menu to insert, delete, hide and unhide, and the
view zooms from 10% to 400%. Each change is a step to undo. Widths, heights and hidden rows and
columns are the worksheet's own, so `XlsxWriter` saves them. See
[What the grid shows](guides/showing-workbooks.md#rows-columns-and-filling) for every key.

## Selection, copying and scrolling

```csharp snippet=active-cell
view.ActiveCell = new CellPosition(row: 10, column: 3);                // C10, selected alone
view.Selection = view.Selection!.Extend(rows: 2, columns: 1, sheet);   // C10:D12, as Shift with the arrows
view.Offset = new Vector(0, 200);                                      // scrolled ten default rows down
await view.CopyAsync();                                                // as Ctrl+C
```

`ActiveCell`, `Selection` and `Offset` are Avalonia properties, so they can be bound. The active
cell is the selection's; setting it selects that cell alone. Cells are selected with the mouse and
keyboard as in Excel, and copying puts them on the clipboard as text and as an HTML table, which
Excel pastes with their styles. The wheel scrolls three rows a turn.
