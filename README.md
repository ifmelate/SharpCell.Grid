# SharpCell.Grid

[![NuGet](https://img.shields.io/nuget/v/SharpCell.Grid.Avalonia.svg)](https://www.nuget.org/packages/SharpCell.Grid.Avalonia)
[![ci](https://github.com/ifmelate/SharpCell.Grid/actions/workflows/ci.yml/badge.svg)](https://github.com/ifmelate/SharpCell.Grid/actions/workflows/ci.yml)

A free (MIT) Excel-like spreadsheet control for .NET 8 and later, built on the
[SharpCell](https://github.com/ifmelate/SharpCell) formula engine. It shows a worksheet the way
Excel does: column and row headers, gridlines, each cell's value in its number format and its
style (font, fill, borders, alignment, wrapping), the file's column widths and row heights,
merged cells, frozen panes and text running on over empty neighbours. Cells are selected, typed
into, deleted, pasted and undone with the mouse and keyboard as in Excel, the workbook recalculated
after each change, and copied with their styles, ready to paste into Excel. Only the visible cells are read, so a sheet of a million rows scrolls like an empty one.

**[Try it in your browser](https://ifmelate.github.io/SharpCell.Grid/demo/)**: open your own .xlsx file; it is read on your computer and never uploaded.

Status: 0.x, viewing, selecting, copying to Excel and editing. See the [roadmap](docs/roadmap.md).
The API may change between minor versions until 1.0.

| Package | What it does |
|---|---|
| `SharpCell.Grid` | Geometry, hit testing, scrolling and painting, with no UI framework. |
| `SharpCell.Grid.Avalonia` | `SpreadsheetView`, the control for [Avalonia UI](https://avaloniaui.net). |
| `SharpCell.Grid.Wpf` | `SpreadsheetView`, the same control for WPF on Windows. |

## Install

```bash
dotnet add package SharpCell.Grid.Avalonia   # the control; brings SharpCell.Grid and SharpCell
dotnet add package SharpCell.Xlsx            # to read .xlsx files
```

For WPF, add `SharpCell.Grid.Wpf` instead of `SharpCell.Grid.Avalonia`.

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
