# Painting in other UI frameworks

`SharpCell.Grid` holds everything about the grid that is not Avalonia: sizes and positions,
hit testing, scrolling and painting. `SpreadsheetView` is a thin host over it, and a host for
another UI framework is the same few calls.

```csharp snippet=painter
var painter = new GridPainter
{
    // Text width in your UI framework's font engine, for fitting numbers and running text.
    MeasureText = (text, font) => text.Length * font.Size * 0.5,
};
painter.Layout = SheetLayout.FromWorksheet(sheet, painter.MaxDigitWidth(sheet));
painter.FrozenRows = sheet.FrozenRows;
painter.FrozenColumns = sheet.FrozenColumns;

var viewport = new GridViewport(OffsetX: 0, OffsetY: 0, Width: 800, Height: 600);
painter.Paint(canvas, sheet, viewport, activeCell: new CellPosition(1, 1));   // canvas: your IGridCanvas
CellPosition? clicked = painter.HitTest(viewport, x: 120, y: 30, sheet);
```

## The parts

- `SheetLayout` holds the column widths and row heights in pixels, their offsets, the cell under
  an offset and the cells in a rectangle. `SheetLayout.FromWorksheet` converts a sheet's sizes as
  Excel does, from the width of the widest digit of the default font.
- `GridViewport` is what is shown: the scroll offset into the sheet and the size of the grid on
  screen, headers included.
- `GridPainter` paints the part of the sheet in a viewport onto an `IGridCanvas`, and answers where
  a cell is (`CellBounds`), what is under a point (`HitTest`), where to scroll (`ScrollIntoView`,
  `Clamp`), and where the active cell moves (`GridNavigation.Move`).
- `IGridCanvas` is what a host implements over its drawing API: fill and stroke rectangles, draw
  lines with a dash pattern, draw text inside bounds, cut off at them, on one line or wrapped.

## Measuring text

`GridPainter.MeasureText` gives the width of text in a font, from the host's font engine. With it,
numbers are fitted to their column, text runs on over empty cells, and row headers widen for long
row numbers. Without it, nothing is fitted and Calibri 11's digit width (7 pixels) is assumed.

`GridFont` sizes are in device-independent pixels: 11 points are 14.67 pixels at 96 DPI.
`GridPainter.FontOf` converts a SharpCell font.
