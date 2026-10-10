# Working on SharpCell.Grid

Instructions for AI agents and people changing this repository.

## Layout

- `src/SharpCell.Grid` — the grid without a UI framework: row and column geometry (`SheetLayout`),
  hit testing, scrolling and painting through `IGridCanvas` (`GridPainter`). Depends on SharpCell only.
- `src/SharpCell.Grid.Avalonia` — `SpreadsheetView`, the Avalonia control: an `IGridCanvas` over
  Avalonia's drawing context, mouse, keyboard and wheel.
- `tests/SharpCell.Grid.Tests` — xUnit tests: geometry and painting against a recording canvas,
  the Avalonia control headless with rendered frames compared to approved PNGs.
- `samples/SharpCell.Grid.Sample` — a desktop window showing an .xlsx file.

The formula engine and the .xlsx reader are the NuGet packages `SharpCell` and `SharpCell.Xlsx`
from https://github.com/ifmelate/SharpCell. Anything the grid needs from them goes through their
public API; a missing piece is added there, not worked around here.

## Commands

```bash
dotnet build
dotnet test                                                              # net8.0 and net10.0
dotnet run --project samples/SharpCell.Grid.Sample -f net10.0 -- file.xlsx
```

## Snapshots

Tests in `tests/SharpCell.Grid.Tests/Controls` render the control headless with Skia and the
bundled Inter font and compare the frame with `Controls/Snapshots/<name>.png`. A missing or
different snapshot fails and writes `<name>.received.png` next to it. Look at the received image;
when it is right, rename it to `<name>.png` and commit it with the change that caused it.

## Releases

The version comes from the git tag (MinVer). Pushing a tag `vX.Y.Z` on `main` runs
`.github/workflows/release.yml`: test, pack, publish to nuget.org through Trusted Publishing,
GitHub release. Before 1.0, a minor version may break the API; move `PublicAPI.Unshipped.txt`
entries to `PublicAPI.Shipped.txt` when releasing.

## Rules

- Test first: a failing test, then the code, then the commit.
- Every public API change updates `PublicAPI.Unshipped.txt` and has an XML comment; the build fails otherwise.
- What the grid shows follows Microsoft Excel: cell text comes from SharpCell's `Cell.Text`, and
  geometry and behaviour are checked against Excel, never against other spreadsheet programs.
- `SharpCell.Grid` takes no dependency besides SharpCell and knows no UI framework.
- Code, comments, docs and commit messages are in English. Commits follow `feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `build:`.
