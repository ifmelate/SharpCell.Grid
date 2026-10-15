using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpCell.Grid;

// What a grid needs beyond painting cells, to change them with the mouse as in Excel: the edges of
// headers to drag, the fill handle, the cells copied and filled, and the sizes that fit a column's
// or row's content.
public sealed partial class GridPainter
{
    // How near an edge between headers the mouse grabs it, in pixels on either side.
    private const double EdgeReach = 3;

    // The fill handle: a square of five pixels on the selection's bottom-right corner, in a pixel of
    // background, grabbed two pixels around.
    private const double HandleSize = 5;
    private const double HandleReach = 2;

    // Excel's line height for a font, which an auto-fitted row takes: about 1.32 times its size.
    private const double LineSpacing = 1.32;

    /// <summary>
    /// Whether the selection's bottom-right corner has Excel's fill handle, the square dragged to
    /// fill cells. False by default; a view that changes cells turns it on.
    /// </summary>
    public bool ShowsFillHandle { get; set; }

    /// <summary>The cells copied or cut, painted with a dashed border as Excel shows them until they are pasted; null for none.</summary>
    public CellArea? CopiedArea { get; set; }

    /// <summary>The cells picked with the mouse while a formula is typed, painted with a dashed blue border; null for none.</summary>
    public CellArea? ReferenceArea { get; set; }

    /// <summary>The cells a fill handle is being dragged over, painted with a dashed grey border; null for none.</summary>
    public CellArea? FillArea { get; set; }

    /// <summary>
    /// The column whose right edge is under a point of the column headers, as Excel lets it be
    /// dragged to resize the column; null elsewhere. Hidden columns are never found.
    /// </summary>
    public int? ColumnEdgeAt(GridViewport viewport, double x, double y)
    {
        var rowHeaders = RowHeaders(viewport);
        if (y < 0 || y >= HeadersHeight || x < rowHeaders - EdgeReach || x >= viewport.Width)
            return null;
        var cellX = Math.Max(0, x - rowHeaders);
        var column = ColumnUnder(viewport, cellX);
        var left = ScreenX(viewport, rowHeaders, column);
        if (left + Layout.ColumnWidth(column) - x <= EdgeReach)
            return column;
        if (x - left >= EdgeReach || left - rowHeaders < 1)
            return null;

        // Just right of an edge: the column before it, past any hidden ones.
        var before = ColumnUnder(viewport, left - rowHeaders - 0.5);
        return before < column && Layout.ColumnWidth(before) > 0 ? before : null;
    }

    /// <summary>The row whose bottom edge is under a point of the row headers; null elsewhere. See <see cref="ColumnEdgeAt"/>.</summary>
    public int? RowEdgeAt(GridViewport viewport, double x, double y)
    {
        var rowHeaders = RowHeaders(viewport);
        if (x < 0 || x >= rowHeaders || y < HeadersHeight - EdgeReach || y >= viewport.Height)
            return null;
        var cellY = Math.Max(0, y - HeadersHeight);
        var row = RowUnder(viewport, cellY);
        var top = ScreenY(viewport, row);
        if (top + Layout.RowHeight(row) - y <= EdgeReach)
            return row;
        if (y - top >= EdgeReach || top - HeadersHeight < 1)
            return null;
        var before = RowUnder(viewport, top - HeadersHeight - 0.5);
        return before < row && Layout.RowHeight(before) > 0 ? before : null;
    }

    /// <summary>Whether a point is on the fill handle of the selection, when <see cref="ShowsFillHandle"/> shows one.</summary>
    public bool IsOnFillHandle(GridViewport viewport, double x, double y, GridSelection? selection, Worksheet? sheet)
    {
        if (!ShowsFillHandle || selection is null)
            return false;
        var rowHeaders = RowHeaders(viewport);
        var area = GridSelection.Cover(selection.Area, sheet?.MergedAreas ?? []);
        if (FillHandle(viewport, rowHeaders, Regions(viewport, rowHeaders), area) is not { } square)
            return false;
        var grab = square.Handle;
        return x >= grab.X - HandleReach && x < grab.Right + HandleReach && y >= grab.Y - HandleReach && y < grab.Bottom + HandleReach;
    }

    // The fill handle of a selection, when its bottom-right corner is in view: the square and the
    // pixel of background around it.
    private (GridRect Handle, GridRect Border)? FillHandle(GridViewport viewport, double rowHeaders, List<Region> regions, CellArea area)
    {
        var corner = new CellPosition(area.LastRow, area.LastColumn);
        if (regions.FirstOrDefault(r => r.Area.Contains(corner.Row, corner.Column)) is not { Clip.Width: > 0 } region)
            return null;
        var bounds = Bounds(viewport, area, rowHeaders);
        var (right, bottom) = (Math.Floor(bounds.Right) - 1, Math.Floor(bounds.Bottom) - 1);
        var half = Math.Floor(HandleSize / 2);
        var handle = new GridRect(right - half, bottom - half, HandleSize, HandleSize);
        var clip = region.Clip;
        if (handle.X >= clip.Right || handle.Y >= clip.Bottom || handle.Right <= clip.X || handle.Bottom <= clip.Y)
            return null;
        return (handle, new GridRect(handle.X - 1, handle.Y - 1, HandleSize + 2, HandleSize + 2));
    }

    /// <summary>
    /// The width in pixels at <see cref="Zoom"/> that fits the widest text of a column, as Excel's
    /// AutoFit Column Width does; merged cells are left out. Null when the column holds nothing to
    /// fit or without <see cref="MeasureText"/>.
    /// </summary>
    public double? FitColumnWidth(Worksheet sheet, int column)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (MeasureText is not { } measure)
            return null;
        var merges = sheet.MergedAreas;
        var digit = MaxDigitWidth(sheet);
        double? widest = null;
        foreach (var cell in GridEditor.RangeOf(sheet, new CellArea(1, column, SheetLayout.MaxRow, column)).Cells)
        {
            var text = cell.Text;
            if (text.Length == 0 || GridNavigation.MergeAt(merges, new CellPosition(cell.Row, cell.Column)) is not null)
                continue;
            var style = cell.Style;
            var font = Font(style.Font);
            var lines = text.Split('\n').Max(line => measure(line, font));
            var needed = lines + 2 * Padding + 1 + style.Indent * IndentStep * digit / DefaultDigitWidth;
            widest = Math.Max(widest ?? 0, Math.Ceiling(needed));
        }

        return widest;
    }

    /// <summary>
    /// The height in pixels at <see cref="Zoom"/> that fits the tallest text of a row, as Excel's
    /// AutoFit Row Height does: a line of each cell's font, and as many lines as wrapped text or line
    /// breaks take. Null when the row holds nothing, for the sheet's default height.
    /// </summary>
    public double? FitRowHeight(Worksheet sheet, int row)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        var merges = sheet.MergedAreas;
        double? tallest = null;
        foreach (var cell in GridEditor.RangeOf(sheet, new CellArea(row, 1, row, SheetLayout.MaxColumn)).Cells)
        {
            var text = cell.Text;
            if (text.Length == 0 || GridNavigation.MergeAt(merges, new CellPosition(cell.Row, cell.Column)) is not null)
                continue;
            var style = cell.Style;
            var font = Font(style.Font);
            var lines = style.WrapText && MeasureText is { } measure
                ? Wrapped(text, Layout.ColumnWidth(cell.Column) - 2 * Padding - 1, t => measure(t, font))
                : text.Split('\n').Length;
            tallest = Math.Max(tallest ?? 0, lines * Math.Ceiling(font.Size * LineSpacing));
        }

        return tallest;
    }

    // The lines text takes when wrapped to a width: words go on while they fit.
    private static int Wrapped(string text, double width, Func<string, double> measure)
    {
        var lines = 0;
        foreach (var paragraph in text.Split('\n'))
        {
            lines++;
            var line = "";
            foreach (var word in paragraph.Split(' '))
            {
                var next = line.Length == 0 ? word : line + " " + word;
                if (line.Length > 0 && measure(next) > width)
                {
                    lines++;
                    line = word;
                }
                else
                {
                    line = next;
                }
            }
        }

        return lines;
    }
}
