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
- `tools/SharpCell.Grid.Docs` — fills code snippets into Markdown, writes `docs/api`, `docs/llms*.txt`, renders the site. Copied from SharpCell's `tools/SharpCell.Docs`; its unit tests stay there.

The formula engine and the .xlsx reader are the NuGet packages `SharpCell` and `SharpCell.Xlsx`
from https://github.com/ifmelate/SharpCell. Anything the grid needs from them goes through their
public API; a missing piece is added there, not worked around here.

## Commands

```bash
dotnet build
dotnet test                                                              # net8.0 and net10.0
dotnet run --project samples/SharpCell.Grid.Sample -f net10.0 -- file.xlsx
dotnet run --project tools/SharpCell.Grid.Docs -f net10.0 -- generate   # after changing docs, snippets or public API
dotnet run --project tools/SharpCell.Grid.Docs -f net10.0 -- check
dotnet run --project tools/SharpCell.Grid.Docs -f net10.0 -- site artifacts/site
```

## Snapshots

Tests in `tests/SharpCell.Grid.Tests/Controls` render the control headless with Skia and the
bundled Inter font and compare the frame with `Controls/Snapshots/<name>.<os>.png`, one per
operating system (`linux`, `macos`, `windows`): Skia draws glyphs differently on each. A missing
or different snapshot fails and writes `<name>.<os>.received.png` next to it. Look at the received
image; when it is right, rename it to `<name>.<os>.png` and commit it with the change that caused
it. CI runs on Linux: when a change alters a frame, take the Linux images from the
`received-snapshots` artifact of the failed run, check them, and commit them too.

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
- Code examples in docs live in tested `.cs` files between `// snippet: name` and `// end-snippet` (in `tests/SharpCell.Grid.Tests/Docs` or `samples`); never edit the body of a `csharp snippet=` block by hand.
- `docs/api`, `docs/llms.txt` and `docs/llms-full.txt` are generated. Pictures for the site go in `docs/images`.
- Code, comments, docs and commit messages are in English. Commits follow `feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `build:`.
