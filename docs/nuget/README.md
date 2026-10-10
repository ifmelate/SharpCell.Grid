# SharpCell.Grid

A free (MIT) Excel-like spreadsheet control for .NET 8 and later, built on the SharpCell formula
engine. It shows a worksheet as Excel does: headers, gridlines, each cell's value in its number
format and Excel's default alignment. Only the visible cells are read, so any sheet size scrolls
smoothly.

| Package | What it does |
|---|---|
| `SharpCell.Grid` | Geometry, hit testing, scrolling and painting, with no UI framework. |
| `SharpCell.Grid.Avalonia` | `SpreadsheetView`, the control for Avalonia UI. Brings `SharpCell.Grid` with it. |

## Show a workbook

```csharp
using SharpCell;
using SharpCell.Grid.Avalonia;
using SharpCell.Xlsx;

Workbook workbook = XlsxReader.Load("budget.xlsx");
workbook.Recalculate();

var view = new SpreadsheetView { Worksheet = workbook["Budget"] };
```

Source, issues and status: https://github.com/ifmelate/SharpCell.Grid
