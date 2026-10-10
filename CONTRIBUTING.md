# Contributing

Thank you for helping. SharpCell.Grid is kept small and faithful to Excel; contributions that keep
it that way are welcome.

## Bugs

Open an issue with the bug form. It needs a small .xlsx that shows the problem and a screenshot
of the same cells in Microsoft Excel. Other spreadsheet programs are not a reference. An issue
without a reproducing file is closed.

## Pull requests

- Start with a failing test: geometry and painting against the recording canvas, the control
  headless; a visible change comes with its approved snapshot.
- A public API change updates `PublicAPI.Unshipped.txt` and has XML documentation.
- `SharpCell.Grid` depends on SharpCell only and knows no UI framework.
- Code, comments and docs are in English.

By contributing you agree that your contribution is licensed under the MIT license.
