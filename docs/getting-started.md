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
