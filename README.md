# SharpCell.Grid

[![ci](https://github.com/ifmelate/SharpCell.Grid/actions/workflows/ci.yml/badge.svg)](https://github.com/ifmelate/SharpCell.Grid/actions/workflows/ci.yml)

A free (MIT) Excel-like spreadsheet control for .NET 8 and later, built on the
[SharpCell](https://github.com/ifmelate/SharpCell) formula engine. It shows a worksheet the way
Excel does: column and row headers, gridlines, each cell's value in its number format and its
style (font, fill, borders, alignment, wrapping), the file's column widths and row heights,
merged cells, frozen panes, text running on over empty neighbours, and an active cell moved with
the mouse and keyboard. Only the visible cells are read, so a sheet of a million rows scrolls like an empty one.

Status: early 0.x, not on NuGet yet; it needs SharpCell 0.6.0, which is not released either.
Viewing only, no editing; see the [roadmap](docs/roadmap.md).

| Package | What it does |
|---|---|
| `SharpCell.Grid` | Geometry, hit testing, scrolling and painting, with no UI framework. |
| `SharpCell.Grid.Avalonia` | `SpreadsheetView`, the control for [Avalonia UI](https://avaloniaui.net). |

## Usage

Namespaces: `SharpCell`, `SharpCell.Xlsx` and `SharpCell.Grid.Avalonia`.

```csharp snippet=show-workbook
Workbook workbook = XlsxReader.Load(path);
workbook.Recalculate();

var view = new SpreadsheetView { Worksheet = workbook["Budget"] };
```

After changing cells or recalculating, call `view.InvalidateVisual()`. A runnable window is in
[samples/SharpCell.Grid.Sample](samples/SharpCell.Grid.Sample).

## Documentation

- [Documentation site](https://ifmelate.github.io/SharpCell.Grid/): getting started, what the grid shows, painting in other UI frameworks, API reference.
- [Roadmap](docs/roadmap.md).
- [For AI coding agents](docs/agents.md): `llms.txt` and `llms-full.txt`.

## Support

SharpCell.Grid is maintained by one person in spare time, with no support obligations. A bug
report needs a small .xlsx that shows the problem and a screenshot of the same cells in Microsoft
Excel; an issue without both is closed. Security issues: see [SECURITY.md](SECURITY.md).
Contributing: see [CONTRIBUTING.md](CONTRIBUTING.md).
