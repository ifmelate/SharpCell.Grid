# SharpCell.Grid

A free (MIT) Excel-like spreadsheet control for .NET 8 and later, built on the SharpCell formula
engine. It shows a worksheet as Excel does: headers, gridlines, each cell's value in its number
format and its style, column widths, row heights, merged cells and frozen panes. Cells are
selected as in Excel and copied with their styles, ready to paste into Excel. Only the visible
cells are read, so any sheet size scrolls smoothly.

| Package | What it does |
|---|---|
| `SharpCell.Grid` | Geometry, hit testing, scrolling and painting, with no UI framework. |
| `SharpCell.Grid.Avalonia` | `SpreadsheetView`, the control for Avalonia UI. Brings `SharpCell.Grid` with it. |
| `SharpCell.Grid.Wpf` | `SpreadsheetView`, the same control for WPF on Windows. Brings `SharpCell.Grid` with it. |

## Show a workbook

Namespaces: `SharpCell`, `SharpCell.Xlsx` and `SharpCell.Grid.Avalonia`.

```csharp snippet=show-workbook
Workbook workbook = XlsxReader.Load(path);
workbook.Recalculate();

var view = new SpreadsheetView { Worksheet = workbook["Budget"] };
```

Documentation: https://ifmelate.github.io/SharpCell.Grid/ · source and issues: https://github.com/ifmelate/SharpCell.Grid
