# Getting started

## Install

```bash
dotnet add package SharpCell.Grid.Avalonia   # the control; brings SharpCell.Grid and SharpCell
dotnet add package SharpCell.Xlsx            # to read .xlsx files
```

The packages target .NET 8 and .NET 10 and Avalonia 12.

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

## The active cell and scrolling

```csharp snippet=active-cell
view.ActiveCell = new CellPosition(row: 10, column: 3);   // C10
view.Offset = new Vector(0, 200);                // scrolled ten default rows down
```

`ActiveCell` and `Offset` are Avalonia properties, so they can be bound. The mouse picks the
active cell, the arrow keys, Enter and Tab move it as in Excel, and the wheel scrolls three rows a
turn.
