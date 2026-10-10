using System;
using System.Collections.Generic;
using System.Globalization;

namespace SharpCell.Grid;

/// <summary>
/// Paints a worksheet the way Excel shows it: column and row headers, gridlines, each cell's
/// <see cref="Cell.Text"/> in its style (font, fill, borders, alignment, wrapping, indent), merged
/// areas as one cell, and the active cell. Only the cells inside the viewport are read, so a
/// sheet's size does not matter. Also answers the geometry questions of a grid on screen: where a
/// cell is, what is under a point, where to scroll. With <see cref="MeasureText"/> set, text too
/// wide for its cell runs on over empty neighbours, numbers too wide for their column are shortened
/// or shown as <c>###</c>, and row headers widen for long row numbers, as in Excel.
/// </summary>
public sealed class GridPainter
{
    // Space between a cell's edges and its text.
    private const double TextPadding = 2;
    private const double ActiveBorderThickness = 2;

    // Space on each side of a row number, and the digits row headers always have room for.
    private const double HeaderPadding = 6;
    private const int MinRowHeaderDigits = 3;

    // Calibri 11's widest digit, which Excel's sizes are based on, when no text can be measured.
    private const double DefaultDigitWidth = 7;

    // An indent step of Excel at 100% zoom with Calibri 11, in pixels.
    private const double IndentStep = 9;

    // How far left or right of the viewport a cell is looked for whose text runs into view.
    private const int OverflowReach = 128;

    private SheetLayout _layout;

    /// <summary>A painter with Excel's default row heights and column widths.</summary>
    public GridPainter()
        : this(new SheetLayout())
    {
    }

    /// <summary>A painter over the given sizes of rows and columns.</summary>
    public GridPainter(SheetLayout layout)
    {
        ArgumentNullException.ThrowIfNull(layout);
        _layout = layout;
    }

    /// <summary>The sizes of rows and columns; <see cref="SheetLayout.FromWorksheet"/> gives a sheet's own.</summary>
    public SheetLayout Layout
    {
        get => _layout;
        set => _layout = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>The colours of everything but the cells' own styles.</summary>
    public GridTheme Theme { get; set; } = GridTheme.Excel;

    private int _frozenRows;
    private int _frozenColumns;

    /// <summary>
    /// How many rows at the top stay in place while the rest scrolls, as Excel's Freeze Panes;
    /// a sheet's own is <see cref="Worksheet.FrozenRows"/>.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The count is negative or leaves no row to scroll.</exception>
    public int FrozenRows
    {
        get => _frozenRows;
        set => _frozenRows = value is >= 0 and < SheetLayout.MaxRow ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "Frozen rows must leave a row to scroll.");
    }

    /// <summary>How many columns on the left stay in place while the rest scrolls; a sheet's own is <see cref="Worksheet.FrozenColumns"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The count is negative or leaves no column to scroll.</exception>
    public int FrozenColumns
    {
        get => _frozenColumns;
        set => _frozenColumns = value is >= 0 and < SheetLayout.MaxColumn ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "Frozen columns must leave a column to scroll.");
    }

    /// <summary>The height of the column headers at the top.</summary>
    public double ColumnHeaderHeight { get; set; } = 20;

    /// <summary>
    /// The width of the row headers on the left; with <see cref="MeasureText"/>, the least width,
    /// widened to fit the number of the last visible row.
    /// </summary>
    public double RowHeaderWidth { get; set; } = 40;

    /// <summary>
    /// The width in pixels of a line of text in a font. Without it, text is never fitted to its
    /// cell, row headers keep <see cref="RowHeaderWidth"/>, and Calibri 11's digit width is assumed.
    /// </summary>
    public Func<string, GridFont, double>? MeasureText { get; set; }

    /// <summary>The font of a SharpCell font at 96 DPI.</summary>
    public static GridFont FontOf(CellFont font)
    {
        ArgumentNullException.ThrowIfNull(font);
        return new GridFont(font.Name, font.Size * 96 / 72, font.Bold, font.Italic, font.Underline != CellUnderline.None, font.Strikethrough);
    }

    /// <summary>
    /// The width in pixels of the widest digit of a sheet's default font, which Excel bases column
    /// widths on; 7, Calibri 11's, without <see cref="MeasureText"/>.
    /// </summary>
    public double MaxDigitWidth(Worksheet? sheet)
    {
        if (MeasureText is not { } measure)
            return DefaultDigitWidth;
        var font = DefaultFont(sheet);
        var widest = 0.0;
        for (var digit = '0'; digit <= '9'; digit++)
            widest = Math.Max(widest, measure(digit.ToString(), font));
        return widest > 0 ? Math.Round(widest) : DefaultDigitWidth;
    }

    /// <summary>Paints the part of the sheet in the viewport with the active cell alone selected; without a sheet, an empty grid.</summary>
#pragma warning disable RS0027 // The overload taking a selection came after this one shipped; they differ in the last parameter's type.
    public void Paint(IGridCanvas canvas, Worksheet? sheet, GridViewport viewport, CellPosition? activeCell = null) =>
        Paint(canvas, sheet, viewport, activeCell is { } cell ? GridSelection.Cell(cell, sheet) : null);
#pragma warning restore RS0027

    /// <summary>
    /// Paints the part of the sheet in the viewport and the selection over it, as Excel does: the
    /// selected cells but the active one shaded, a border around them, their headers highlighted.
    /// Without a sheet, an empty grid; without a selection, none.
    /// </summary>
    public void Paint(IGridCanvas canvas, Worksheet? sheet, GridViewport viewport, GridSelection? selection)
    {
        ArgumentNullException.ThrowIfNull(canvas);
        var theme = Theme;
        canvas.FillRectangle(new GridRect(0, 0, viewport.Width, viewport.Height), theme.Background);
        if (CellsWidth(viewport) <= 0 || CellsHeight(viewport) <= 0)
            return;

        // The selection covers merged areas whole, also when it was made without the sheet.
        var merges = sheet?.MergedAreas ?? [];
        Selected? selected = selection is null ? null : new Selected(
            GridSelection.Cover(selection.Area, merges),
            GridSelection.Cover(GridSelection.Box(selection.ActiveCell, selection.ActiveCell), merges),
            selection.IsEntireColumns,
            selection.IsEntireRows);

        // The scrolled cells first, then the frozen panes over whatever of them reaches under the
        // panes: each frozen pane starts from its own background.
        var rowHeaders = RowHeaders(viewport);
        var regions = Regions(viewport, rowHeaders);
        var frames = new List<Frame>();
        for (var i = 0; i < regions.Count; i++)
        {
            var frame = new Frame(this, sheet, viewport, rowHeaders, regions[i]);
            frames.Add(frame);
            if (i > 0)
                canvas.FillRectangle(regions[i].Clip, theme.Background);
            if (sheet?.ShowGridlines ?? true)
                PaintGridlines(canvas, frame);
            if (sheet is not null)
            {
                PaintFills(canvas, frame, sheet);
                var runs = TextRuns(frame, sheet);
                CoverGridlines(canvas, frame, sheet, runs);
                PaintBorders(canvas, frame, sheet);
                foreach (var run in runs)
                    PaintText(canvas, frame, sheet, run);
            }

            if (selected is { } shade)
                ShadeSelection(canvas, frame, shade);
        }

        PaintFrozenLines(canvas, viewport, rowHeaders);
        PaintHeaders(canvas, frames, selected);
        if (selected is { } outline)
        {
            foreach (var frame in frames)
                OutlineSelection(canvas, frame, outline.Area);
        }
    }

    /// <summary>A selection as painted: its cells and the active cell's, merged areas covered.</summary>
    private readonly record struct Selected(CellArea Area, CellArea Active, bool EntireColumns, bool EntireRows);

    // Excel shades the selected cells, text and all, but leaves the active cell white.
    private void ShadeSelection(IGridCanvas canvas, Frame frame, Selected selected)
    {
        if (selected.Area == selected.Active || !Intersects(frame.Area, selected.Area))
            return;
        var area = frame.Bounds(selected.Area);
        var active = frame.Bounds(selected.Active);
        var clip = frame.Clip;
        void Fill(double left, double top, double right, double bottom)
        {
            left = Math.Max(left, clip.X);
            top = Math.Max(top, clip.Y);
            right = Math.Min(right, clip.Right);
            bottom = Math.Min(bottom, clip.Bottom);
            if (right > left && bottom > top)
                canvas.FillRectangle(new GridRect(left, top, right - left, bottom - top), Theme.SelectionFill);
        }

        // Around the active cell: the bands above and below it, then those beside it.
        Fill(area.X, area.Y, area.Right, active.Y);
        Fill(area.X, active.Bottom, area.Right, area.Bottom);
        Fill(area.X, Math.Max(area.Y, active.Y), active.X, Math.Min(area.Bottom, active.Bottom));
        Fill(active.Right, Math.Max(area.Y, active.Y), area.Right, Math.Min(area.Bottom, active.Bottom));
    }

    // The border around the selection, two whole pixels wide: the gridline on each side and the
    // pixel inside it. It is drawn as the part of it each pane shows: a selection reaching under a
    // frozen pane is cut off at its line. Edges along the headers lie over the headers' closing line.
    private void OutlineSelection(IGridCanvas canvas, Frame frame, CellArea selection)
    {
        if (!Intersects(frame.Area, selection))
            return;
        var bounds = frame.Bounds(selection);
        var clip = frame.Clip;
        var color = Theme.ActiveCellBorder;
        var (left, top) = (Math.Floor(bounds.X), Math.Floor(bounds.Y));
        var (right, bottom) = (Math.Floor(bounds.Right) - 1, Math.Floor(bounds.Bottom) - 1);
        var half = ActiveBorderThickness / 2;

        var (y1, y2) = (Math.Max(top - half, clip.Y - half), Math.Min(bottom + half, clip.Bottom));
        if (y2 > y1)
        {
            if (left >= clip.X - half && left < clip.Right)
                canvas.DrawLine(left, y1, left, y2, color, ActiveBorderThickness);
            if (right >= clip.X - half && right < clip.Right)
                canvas.DrawLine(right, y1, right, y2, color, ActiveBorderThickness);
        }

        var (x1, x2) = (Math.Max(left - half, clip.X - half), Math.Min(right + half, clip.Right));
        if (x2 > x1)
        {
            if (top >= clip.Y - half && top < clip.Bottom)
                canvas.DrawLine(x1, top, x2, top, color, ActiveBorderThickness);
            if (bottom >= clip.Y - half && bottom < clip.Bottom)
                canvas.DrawLine(x1, bottom, x2, bottom, color, ActiveBorderThickness);
        }
    }

    /// <summary>The cells the viewport shows, partly or wholly, of the part that scrolls: frozen rows and columns aside.</summary>
    public CellArea VisibleArea(GridViewport viewport)
    {
        var (columns, rows) = (ScrollColumns(viewport), ScrollRows(viewport));
        return new CellArea(rows.First, columns.First, rows.Last, columns.Last);
    }

    // The columns of the part that scrolls in view, and those of the frozen part; a range may be
    // empty (Last < First) when the frozen part fills the view.
    private (int First, int Last) ScrollColumns(GridViewport viewport)
    {
        var width = Math.Max(0, CellsWidth(viewport));
        var offset = Math.Max(0, viewport.OffsetX);
        var frozen = Math.Min(FrozenWidth, width);
        var first = Math.Max(FrozenColumns + 1, Layout.ColumnAt(frozen + offset));
        var last = width <= frozen ? first - 1 : Layout.ColumnAt(offset + Math.BitDecrement(width));
        return (first, last);
    }

    private (int First, int Last) ScrollRows(GridViewport viewport)
    {
        var height = Math.Max(0, CellsHeight(viewport));
        var offset = Math.Max(0, viewport.OffsetY);
        var frozen = Math.Min(FrozenHeight, height);
        var first = Math.Max(FrozenRows + 1, Layout.RowAt(frozen + offset));
        var last = height <= frozen ? first - 1 : Layout.RowAt(offset + Math.BitDecrement(height));
        return (first, last);
    }

    private (int First, int Last) FrozenColumnRange(GridViewport viewport)
    {
        var width = Math.Max(0, CellsWidth(viewport));
        if (FrozenColumns == 0 || width == 0)
            return (1, 0);
        return (1, Math.Min(FrozenColumns, Layout.ColumnAt(Math.BitDecrement(Math.Min(FrozenWidth, width)))));
    }

    private (int First, int Last) FrozenRowRange(GridViewport viewport)
    {
        var height = Math.Max(0, CellsHeight(viewport));
        if (FrozenRows == 0 || height == 0)
            return (1, 0);
        return (1, Math.Min(FrozenRows, Layout.RowAt(Math.BitDecrement(Math.Min(FrozenHeight, height)))));
    }

    private double FrozenWidth => FrozenColumns == 0 ? 0 : Layout.ColumnLeft(FrozenColumns) + Layout.ColumnWidth(FrozenColumns);

    private double FrozenHeight => FrozenRows == 0 ? 0 : Layout.RowTop(FrozenRows) + Layout.RowHeight(FrozenRows);

    /// <summary>A part of the grid painted on its own: the cells in it, the columns its text may run over, its place on screen.</summary>
    private readonly record struct Region(CellArea Area, int MinColumn, int MaxColumn, GridRect Clip);

    // The part that scrolls first, then the frozen columns, the frozen rows and their corner.
    private List<Region> Regions(GridViewport viewport, double rowHeaders)
    {
        var (scrollColumns, scrollRows) = (ScrollColumns(viewport), ScrollRows(viewport));
        var (frozenColumns, frozenRows) = (FrozenColumnRange(viewport), FrozenRowRange(viewport));
        var frozenWidth = Math.Min(FrozenWidth, Math.Max(0, CellsWidth(viewport)));
        var frozenHeight = Math.Min(FrozenHeight, Math.Max(0, CellsHeight(viewport)));
        var left = new GridRect(rowHeaders, 0, frozenWidth, 0);
        var right = new GridRect(rowHeaders + frozenWidth, 0, Math.Max(0, viewport.Width - rowHeaders - frozenWidth), 0);
        var top = (Y: ColumnHeaderHeight, Height: frozenHeight);
        var bottom = (Y: ColumnHeaderHeight + frozenHeight, Height: Math.Max(0, viewport.Height - ColumnHeaderHeight - frozenHeight));

        var regions = new List<Region>();
        void Add((int First, int Last) rows, (int First, int Last) columns, GridRect x, (double Y, double Height) y, bool frozenColumnsPart)
        {
            if (rows.Last < rows.First || columns.Last < columns.First)
                return;
            regions.Add(new Region(
                new CellArea(rows.First, columns.First, rows.Last, columns.Last),
                frozenColumnsPart ? 1 : FrozenColumns + 1,
                frozenColumnsPart ? FrozenColumns : SheetLayout.MaxColumn,
                new GridRect(x.X, y.Y, x.Width, y.Height)));
        }

        Add(scrollRows, scrollColumns, right, bottom, false);
        Add(scrollRows, frozenColumns, left, bottom, true);
        Add(frozenRows, scrollColumns, right, top, false);
        Add(frozenRows, frozenColumns, left, top, true);
        return regions;
    }

    // Columns left of the frozen line do not scroll; the rest move by the offset.
    private double ScreenX(GridViewport viewport, double rowHeaders, int column) =>
        rowHeaders + Layout.ColumnLeft(column) - (column > FrozenColumns ? viewport.OffsetX : 0);

    private double ScreenY(GridViewport viewport, int row) =>
        ColumnHeaderHeight + Layout.RowTop(row) - (row > FrozenRows ? viewport.OffsetY : 0);

    private void PaintFrozenLines(IGridCanvas canvas, GridViewport viewport, double rowHeaders)
    {
        if (FrozenColumns > 0 && FrozenWidth < CellsWidth(viewport))
        {
            var x = PixelCentre(rowHeaders + FrozenWidth);
            canvas.DrawLine(x, ColumnHeaderHeight, x, viewport.Height, Theme.FrozenLine);
        }

        if (FrozenRows > 0 && FrozenHeight < CellsHeight(viewport))
        {
            var y = PixelCentre(ColumnHeaderHeight + FrozenHeight);
            canvas.DrawLine(rowHeaders, y, viewport.Width, y, Theme.FrozenLine);
        }
    }

    /// <summary>Where a cell is on screen; it may lie outside the viewport.</summary>
    public GridRect CellBounds(GridViewport viewport, CellPosition cell) => CellBounds(viewport, cell, cell);

    /// <summary>Where a rectangle of cells is on screen, from its top-left to its bottom-right cell.</summary>
    public GridRect CellBounds(GridViewport viewport, CellPosition first, CellPosition last) =>
        Bounds(viewport, new CellArea(first.Row, first.Column, last.Row, last.Column), RowHeaders(viewport));

    /// <summary>
    /// The cell under a point of the grid; null on the headers or outside the grid. Inside a
    /// merged area of the sheet, its top-left cell.
    /// </summary>
    public CellPosition? HitTest(GridViewport viewport, double x, double y, Worksheet? sheet = null)
    {
        var rowHeaders = RowHeaders(viewport);
        if (x < rowHeaders || y < ColumnHeaderHeight || x >= viewport.Width || y >= viewport.Height)
            return null;
        var (cellX, cellY) = (x - rowHeaders, y - ColumnHeaderHeight);
        var cell = new CellPosition(
            Layout.RowAt(cellY < FrozenHeight ? cellY : cellY + viewport.OffsetY),
            Layout.ColumnAt(cellX < FrozenWidth ? cellX : cellX + viewport.OffsetX));
        return sheet is not null && GridNavigation.MergeAt(sheet.MergedAreas, cell) is { } merge
            ? new CellPosition(merge.FirstRow, merge.FirstColumn)
            : cell;
    }

    /// <summary>
    /// What is under a point of the grid: a cell, a column or row header, or the corner above the
    /// row headers; null outside the grid. <see cref="GridSelection.At"/> turns it into what a click selects.
    /// </summary>
    public GridHit? Hit(GridViewport viewport, double x, double y, Worksheet? sheet = null)
    {
        var rowHeaders = RowHeaders(viewport);
        if (x < 0 || y < 0 || x >= viewport.Width || y >= viewport.Height)
            return null;
        var (onRowHeaders, onColumnHeaders) = (x < rowHeaders, y < ColumnHeaderHeight);
        if (!onRowHeaders && !onColumnHeaders)
            return HitTest(viewport, x, y, sheet) is { } cell ? new GridHit(GridPart.Cell, cell) : null;

        // On a header, the cell a click makes active: the first on screen in that column or row.
        var row = onColumnHeaders ? TopRow(viewport) : RowUnder(viewport, y - ColumnHeaderHeight);
        var column = onRowHeaders ? LeftColumn(viewport) : ColumnUnder(viewport, x - rowHeaders);
        var part = onRowHeaders && onColumnHeaders ? GridPart.Corner : onColumnHeaders ? GridPart.ColumnHeader : GridPart.RowHeader;
        return new GridHit(part, new CellPosition(row, column));
    }

    /// <summary>
    /// The cell nearest to a point, for dragging a selection: the cell under it, or past an edge of
    /// the cells the next one out of view that way, so that scrolling to it follows the mouse.
    /// </summary>
    public CellPosition CellNearest(GridViewport viewport, double x, double y)
    {
        var rowHeaders = RowHeaders(viewport);
        var (width, height) = (Math.Max(0, CellsWidth(viewport)), Math.Max(0, CellsHeight(viewport)));
        var cellX = Math.Clamp(x - rowHeaders, 0, width == 0 ? 0 : Math.BitDecrement(width));
        var cellY = Math.Clamp(y - ColumnHeaderHeight, 0, height == 0 ? 0 : Math.BitDecrement(height));
        var row = RowUnder(viewport, cellY);
        var column = ColumnUnder(viewport, cellX);
        if (y >= viewport.Height)
            row = ScrollRows(viewport).Last + 1;
        else if (y < ColumnHeaderHeight && viewport.OffsetY > 0)
            row = ScrollRows(viewport).First - 1;
        if (x >= viewport.Width)
            column = ScrollColumns(viewport).Last + 1;
        else if (x < rowHeaders && viewport.OffsetX > 0)
            column = ScrollColumns(viewport).First - 1;
        return new CellPosition(Math.Clamp(row, 1, SheetLayout.MaxRow), Math.Clamp(column, 1, SheetLayout.MaxColumn));
    }

    // The row and column under a point of the cells, frozen panes included.
    private int RowUnder(GridViewport viewport, double cellY) => Layout.RowAt(cellY < FrozenHeight ? cellY : cellY + viewport.OffsetY);

    private int ColumnUnder(GridViewport viewport, double cellX) => Layout.ColumnAt(cellX < FrozenWidth ? cellX : cellX + viewport.OffsetX);

    private int TopRow(GridViewport viewport) => FrozenRows > 0 ? 1 : ScrollRows(viewport).First;

    private int LeftColumn(GridViewport viewport) => FrozenColumns > 0 ? 1 : ScrollColumns(viewport).First;

    /// <summary>
    /// The viewport scrolled just enough to show the whole cell, or its top-left part when it is
    /// larger than the viewport; unchanged when the cell is already in view.
    /// </summary>
    public GridViewport ScrollIntoView(GridViewport viewport, CellPosition cell)
    {
        // Frozen cells are always in view; the others scroll in the part after the frozen line.
        var offsetX = cell.Column <= FrozenColumns
            ? viewport.OffsetX
            : Scroll(viewport.OffsetX, CellsWidth(viewport) - FrozenWidth, Layout.ColumnLeft(cell.Column) - FrozenWidth, Layout.ColumnWidth(cell.Column));
        var offsetY = cell.Row <= FrozenRows
            ? viewport.OffsetY
            : Scroll(viewport.OffsetY, CellsHeight(viewport) - FrozenHeight, Layout.RowTop(cell.Row) - FrozenHeight, Layout.RowHeight(cell.Row));
        return viewport with { OffsetX = offsetX, OffsetY = offsetY };
    }

    /// <summary>The viewport with its offsets kept between the start and the end of the sheet.</summary>
    public GridViewport Clamp(GridViewport viewport) => viewport with
    {
        OffsetX = Math.Clamp(viewport.OffsetX, 0, Math.Max(0, Layout.Width - CellsWidth(viewport))),
        OffsetY = Math.Clamp(viewport.OffsetY, 0, Math.Max(0, Layout.Height - CellsHeight(viewport))),
    };

    private static double Scroll(double offset, double visible, double start, double size)
    {
        if (start < offset || size >= visible)
            return start;
        if (start + size > offset + visible)
            return start + size - visible;
        return offset;
    }

    private static bool Intersects(CellArea a, CellArea b) =>
        a.FirstRow <= b.LastRow && b.FirstRow <= a.LastRow && a.FirstColumn <= b.LastColumn && b.FirstColumn <= a.LastColumn;

    private static GridFont DefaultFont(Worksheet? sheet) => FontOf(sheet?.Workbook.DefaultStyle.Font ?? CellFont.Default);

    private static GridColor ColorOf(CellColor color) => new(color.R, color.G, color.B);

    private double CellsWidth(GridViewport viewport) => viewport.Width - RowHeaders(viewport);

    private double CellsHeight(GridViewport viewport) => viewport.Height - ColumnHeaderHeight;

    private GridRect Bounds(GridViewport viewport, CellArea area, double rowHeaders)
    {
        var x = ScreenX(viewport, rowHeaders, area.FirstColumn);
        var y = ScreenY(viewport, area.FirstRow);
        return new GridRect(
            x,
            y,
            ScreenX(viewport, rowHeaders, area.LastColumn) + Layout.ColumnWidth(area.LastColumn) - x,
            ScreenY(viewport, area.LastRow) + Layout.RowHeight(area.LastRow) - y);
    }

    // Excel widens row headers for the digits of the visible row numbers. Rows do not depend on
    // the headers' width, so the last visible row is found first.
    private double RowHeaders(GridViewport viewport)
    {
        if (MeasureText is not { } measure)
            return RowHeaderWidth;
        var height = Math.Max(0, CellsHeight(viewport));
        var lastRow = Math.Max(FrozenRows, Layout.RowAt(Math.Max(0, viewport.OffsetY) + (height == 0 ? 0 : Math.BitDecrement(height))));
        var digits = Math.Max(MinRowHeaderDigits, lastRow.ToString(CultureInfo.InvariantCulture).Length);
        return Math.Max(RowHeaderWidth, Math.Ceiling(measure(new string('0', digits), CellFontForHeaders)) + 2 * HeaderPadding);
    }

    // Headers are measured in Excel's default font: the workbook's is not known outside a paint.
    private static GridFont CellFontForHeaders => FontOf(CellFont.Default);

    private void PaintGridlines(IGridCanvas canvas, Frame frame)
    {
        var color = Theme.Gridline;
        var viewport = frame.Viewport;
        var area = frame.Area;
        var clip = frame.Clip;
        for (var column = area.FirstColumn; column <= area.LastColumn; column++)
        {
            if (Layout.ColumnWidth(column) <= 0)
                continue;
            var right = ScreenX(viewport, frame.RowHeaders, column) + Layout.ColumnWidth(column);
            canvas.DrawLine(PixelCentre(right), clip.Y, PixelCentre(right), clip.Bottom, color);
        }

        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            if (Layout.RowHeight(row) <= 0)
                continue;
            var bottom = ScreenY(viewport, row) + Layout.RowHeight(row);
            canvas.DrawLine(clip.X, PixelCentre(bottom), clip.Right, PixelCentre(bottom), color);
        }
    }

    // Fills lie over the gridlines, as in Excel. A merged area is one cell: its top-left cell's
    // fill covers it whole, and without a fill the background covers the gridlines inside it.
    private void PaintFills(IGridCanvas canvas, Frame frame, Worksheet sheet)
    {
        foreach (var merge in frame.Merges)
        {
            var bounds = frame.Bounds(merge);
            if (sheet[merge.FirstRow, merge.FirstColumn].Style.Fill is { } fill)
                canvas.FillRectangle(bounds, ColorOf(fill));
            else
                canvas.FillRectangle(new GridRect(bounds.X, bounds.Y, Math.Max(0, bounds.Width - 1), Math.Max(0, bounds.Height - 1)), Theme.Background);
        }

        foreach (var (row, column) in frame.Cells())
        {
            if (frame.MergeAt(new CellPosition(row, column)) is not null)
                continue;
            if (sheet[row, column].Style.Fill is { } fill)
                canvas.FillRectangle(frame.Bounds(new CellArea(row, column, row, column)), ColorOf(fill));
        }
    }

    private void PaintBorders(IGridCanvas canvas, Frame frame, Worksheet sheet)
    {
        foreach (var (row, column) in frame.Cells())
        {
            var style = sheet[row, column].Style;
            if (style.LeftBorder is null && style.RightBorder is null && style.TopBorder is null && style.BottomBorder is null)
                continue;

            // Inside a merged area only the edges on its outline show.
            var merge = frame.MergeAt(new CellPosition(row, column));
            var bounds = frame.Bounds(new CellArea(row, column, row, column));
            if (bounds.Width <= 0 || bounds.Height <= 0)
                continue;
            if (style.LeftBorder is { } left && (merge is null || merge.Value.FirstColumn == column))
                Edge(canvas, left, PixelCentre(bounds.X), bounds.Y - 1, PixelCentre(bounds.X), bounds.Bottom, vertical: true);
            if (style.RightBorder is { } right && (merge is null || merge.Value.LastColumn == column))
                Edge(canvas, right, PixelCentre(bounds.Right), bounds.Y - 1, PixelCentre(bounds.Right), bounds.Bottom, vertical: true);
            if (style.TopBorder is { } top && (merge is null || merge.Value.FirstRow == row))
                Edge(canvas, top, bounds.X - 1, PixelCentre(bounds.Y), bounds.Right, PixelCentre(bounds.Y), vertical: false);
            if (style.BottomBorder is { } bottom && (merge is null || merge.Value.LastRow == row))
                Edge(canvas, bottom, bounds.X - 1, PixelCentre(bounds.Bottom), bounds.Right, PixelCentre(bounds.Bottom), vertical: false);
        }
    }

    // A border edge on the pixel line between two cells; thicker lines grow to both sides of it,
    // a double line is two lines with a pixel between them.
    private static void Edge(IGridCanvas canvas, CellBorder border, double x1, double y1, double x2, double y2, bool vertical)
    {
        var color = border.Color is { } c ? ColorOf(c) : new GridColor(0, 0, 0);
        if (border.Style == CellBorderStyle.Double)
        {
            var (dx, dy) = vertical ? (1.0, 0.0) : (0.0, 1.0);
            canvas.DrawLine(x1 - dx, y1 - dy, x2 - dx, y2 - dy, color);
            canvas.DrawLine(x1 + dx, y1 + dy, x2 + dx, y2 + dy, color);
            return;
        }

        var (thickness, dash) = border.Style switch
        {
            CellBorderStyle.Medium => (2.0, GridDash.Solid),
            CellBorderStyle.Thick => (3.0, GridDash.Solid),
            CellBorderStyle.Hair => (1.0, GridDash.Hair),
            CellBorderStyle.Dotted => (1.0, GridDash.Dot),
            CellBorderStyle.Dashed => (1.0, GridDash.Dash),
            CellBorderStyle.DashDot => (1.0, GridDash.DashDot),
            CellBorderStyle.DashDotDot => (1.0, GridDash.DashDotDot),
            CellBorderStyle.MediumDashed => (2.0, GridDash.Dash),
            CellBorderStyle.MediumDashDot or CellBorderStyle.SlantDashDot => (2.0, GridDash.DashDot),
            CellBorderStyle.MediumDashDotDot => (2.0, GridDash.DashDotDot),
            _ => (1.0, GridDash.Solid),
        };

        // A two-pixel line on a pixel centre would blur; it moves half a pixel onto whole pixels.
        var shift = thickness == 2 ? 0.5 : 0;
        if (vertical)
            canvas.DrawLine(x1 + shift, y1, x2 + shift, y2, color, thickness, dash);
        else
            canvas.DrawLine(x1, y1 + shift, x2, y2 + shift, color, thickness, dash);
    }

    /// <summary>A cell's text and the cells it is painted across: the cell, its merged area, or the empty cells it runs on over.</summary>
    private readonly record struct TextRun(int Row, int Column, CellArea Span);

    private List<TextRun> TextRuns(Frame frame, Worksheet sheet)
    {
        var runs = new List<TextRun>();
        foreach (var merge in frame.Merges)
            runs.Add(new TextRun(merge.FirstRow, merge.FirstColumn, merge));

        var area = frame.Area;
        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            if (Layout.RowHeight(row) <= 0)
                continue;

            // Text of a cell left or right of the viewport can run into it.
            if (Outside(frame, sheet, row, area.FirstColumn - 1, -1) is { } fromLeft)
                runs.Add(fromLeft);
            if (Outside(frame, sheet, row, area.LastColumn + 1, 1) is { } fromRight)
                runs.Add(fromRight);

            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
            {
                if (Layout.ColumnWidth(column) <= 0 || frame.MergeAt(new CellPosition(row, column)) is not null)
                    continue;
                if (sheet[row, column].Text.Length > 0)
                    runs.Add(new TextRun(row, column, Span(frame, sheet, row, column)));
            }
        }

        return runs;
    }

    // The nearest cell with content beyond the viewport's edge, when its text runs into view.
    private TextRun? Outside(Frame frame, Worksheet sheet, int row, int start, int step)
    {
        for (var column = start; column >= frame.MinColumn && column <= frame.MaxColumn && Math.Abs(column - start) < OverflowReach; column += step)
        {
            if (!frame.IsFree(sheet, row, column))
            {
                if (frame.MergeAt(row, column, all: true) is not null)
                    return null;
                var span = Span(frame, sheet, row, column);
                return Intersects(span, frame.Area) ? new TextRun(row, column, span) : null;
            }
        }

        return null;
    }

    // Excel lets unwrapped text that is wider than its cell run on over empty cells, away from the
    // side it is aligned to; a cell with content or a merged area stops it. Numbers never run on.
    private CellArea Span(Frame frame, Worksheet sheet, int row, int column)
    {
        var single = new CellArea(row, column, row, column);
        var cell = sheet[row, column];
        if (MeasureText is not { } measure || cell.Value.Kind != CellValueKind.Text)
            return single;
        var style = cell.Style;
        if (style.WrapText || style.HorizontalAlignment is CellHorizontalAlignment.Fill or CellHorizontalAlignment.Justify or CellHorizontalAlignment.Distributed)
            return single;

        var needed = measure(cell.Text, FontOf(style.Font)) + 2 * TextPadding + 1 + Indent(style, frame);
        var own = Layout.ColumnWidth(column);
        if (needed <= own)
            return single;

        var alignment = Horizontal(style.HorizontalAlignment, CellValueKind.Text);
        var (first, last) = (column, column);
        if (alignment == GridTextAlignment.Left)
            last = Extend(frame, sheet, row, column, 1, needed - own);
        else if (alignment == GridTextAlignment.Right)
            first = Extend(frame, sheet, row, column, -1, needed - own);
        else
        {
            first = Extend(frame, sheet, row, column, -1, (needed - own) / 2);
            last = Extend(frame, sheet, row, column, 1, (needed - own) / 2);
        }

        return new CellArea(row, first, row, last);
    }

    // The last column, stepping from a cell, that free cells reach before the room is covered.
    private int Extend(Frame frame, Worksheet sheet, int row, int column, int step, double room)
    {
        var reached = column;
        while (room > 0)
        {
            var next = reached + step;
            if (next < frame.MinColumn || next > frame.MaxColumn || !frame.IsFree(sheet, row, next))
                break;
            reached = next;
            room -= Layout.ColumnWidth(next);
        }

        return reached;
    }

    private double Indent(CellStyle style, Frame frame) => style.Indent * IndentStep * frame.DigitWidth / DefaultDigitWidth;

    // Under running text the gridlines between its cells do not show; a filled cell covers its own.
    private void CoverGridlines(IGridCanvas canvas, Frame frame, Worksheet sheet, List<TextRun> runs)
    {
        foreach (var run in runs)
        {
            if (run.Span.FirstColumn == run.Span.LastColumn || frame.MergeAt(new CellPosition(run.Row, run.Column)) is not null)
                continue;
            for (var column = run.Span.FirstColumn; column < run.Span.LastColumn; column++)
            {
                if (Layout.ColumnWidth(column) <= 0 || sheet[run.Row, column].Style.Fill is not null)
                    continue;
                var bounds = frame.Bounds(new CellArea(run.Row, column, run.Row, column));
                canvas.FillRectangle(new GridRect(Math.Floor(bounds.Right) - 1, bounds.Y, 1, Math.Max(0, bounds.Height - 1)), Theme.Background);
            }
        }
    }

    private void PaintText(IGridCanvas canvas, Frame frame, Worksheet sheet, TextRun run)
    {
        var cell = sheet[run.Row, run.Column];
        var text = cell.Text;
        if (text.Length == 0)
            return;
        var bounds = frame.Bounds(run.Span);
        if (bounds.Width <= 0 || bounds.Height <= 0)
            return;

        var style = cell.Style;
        var value = cell.Value;
        var alignment = Horizontal(style.HorizontalAlignment, value.Kind);

        // The text keeps clear of the gridline on the cell's right and bottom edges; an indent
        // moves it away from the edge it is aligned to.
        var indent = Indent(style, frame);
        var left = bounds.X + TextPadding + (alignment == GridTextAlignment.Left ? indent : 0);
        var right = bounds.Right - TextPadding - 1 - (alignment == GridTextAlignment.Right ? indent : 0);
        var inner = new GridRect(left, bounds.Y, Math.Max(0, right - left), Math.Max(0, bounds.Height - 1));

        var font = FontOf(style.Font);
        if (value.Kind == CellValueKind.Number && !style.WrapText && MeasureText is { } measure)
        {
            var general = cell.NumberFormat.Equals("General", StringComparison.OrdinalIgnoreCase);
            text = TextFitting.FitNumber(text, value.AsNumber(), general, sheet.Workbook.Culture, inner.Width, t => measure(t, font));
        }

        var color = style.Font.Color is { } own ? ColorOf(own) : Theme.CellText;
        canvas.DrawText(text, inner, alignment, Vertical(style.VerticalAlignment), color, font, style.WrapText);
    }

    // Excel's General alignment puts numbers (dates too) right, text left, logical values and
    // errors in the middle. Alignments the canvas cannot draw take the nearest it can.
    private static GridTextAlignment Horizontal(CellHorizontalAlignment alignment, CellValueKind kind) => alignment switch
    {
        CellHorizontalAlignment.Left or CellHorizontalAlignment.Fill or CellHorizontalAlignment.Justify => GridTextAlignment.Left,
        CellHorizontalAlignment.Center or CellHorizontalAlignment.CenterContinuous or CellHorizontalAlignment.Distributed => GridTextAlignment.Center,
        CellHorizontalAlignment.Right => GridTextAlignment.Right,
        _ => kind switch
        {
            CellValueKind.Number => GridTextAlignment.Right,
            CellValueKind.Boolean or CellValueKind.Error => GridTextAlignment.Center,
            _ => GridTextAlignment.Left,
        },
    };

    private static GridVerticalAlignment Vertical(CellVerticalAlignment alignment) => alignment switch
    {
        CellVerticalAlignment.Top => GridVerticalAlignment.Top,
        CellVerticalAlignment.Bottom => GridVerticalAlignment.Bottom,
        _ => GridVerticalAlignment.Center,
    };

    private void PaintHeaders(IGridCanvas canvas, List<Frame> frames, Selected? selected)
    {
        var theme = Theme;
        if (frames.Count == 0)
            return;
        var viewport = frames[0].Viewport;
        var rowHeaders = frames[0].RowHeaders;
        var font = frames[0].HeaderFont;
        canvas.FillRectangle(new GridRect(0, 0, viewport.Width, ColumnHeaderHeight), theme.HeaderBackground);
        canvas.FillRectangle(new GridRect(0, 0, rowHeaders, viewport.Height), theme.HeaderBackground);

        // Scrolled headers first, frozen ones over them, as with the cells.
        var columns = new List<int>();
        var rows = new List<int>();
        foreach (var frame in frames)
        {
            for (var column = frame.Area.FirstColumn; column <= frame.Area.LastColumn; column++)
            {
                if (!columns.Contains(column))
                    columns.Add(column);
            }

            for (var row = frame.Area.FirstRow; row <= frame.Area.LastRow; row++)
            {
                if (!rows.Contains(row))
                    rows.Add(row);
            }
        }

        // Headers of the selected cells' columns and rows are highlighted and underlined on the side
        // of the cells; those of whole selected columns or rows are darker still.
        var underlines = new List<(double From, double To, bool Column)>();
        foreach (var column in columns)
        {
            var width = Layout.ColumnWidth(column);
            if (width <= 0)
                continue;
            var bounds = new GridRect(ScreenX(viewport, rowHeaders, column), 0, width, ColumnHeaderHeight);
            var highlighted = selected is { } s && column >= s.Area.FirstColumn && column <= s.Area.LastColumn;
            var whole = highlighted && selected!.Value.EntireColumns;
            canvas.FillRectangle(bounds, whole ? theme.SelectedHeaderBackground : highlighted ? theme.ActiveHeaderBackground : theme.HeaderBackground);
            canvas.DrawLine(PixelCentre(bounds.Right), 0, PixelCentre(bounds.Right), ColumnHeaderHeight, theme.HeaderLine);
            canvas.DrawText(ColumnNames.Of(column), bounds, GridTextAlignment.Center, GridVerticalAlignment.Center, whole ? theme.SelectedHeaderText : theme.HeaderText, font, wrap: false);
            if (highlighted)
                underlines.Add((bounds.X - 1, bounds.Right, true));
        }

        foreach (var row in rows)
        {
            var height = Layout.RowHeight(row);
            if (height <= 0)
                continue;
            var bounds = new GridRect(0, ScreenY(viewport, row), rowHeaders, height);
            var highlighted = selected is { } s && row >= s.Area.FirstRow && row <= s.Area.LastRow;
            var whole = highlighted && selected!.Value.EntireRows;
            canvas.FillRectangle(bounds, whole ? theme.SelectedHeaderBackground : highlighted ? theme.ActiveHeaderBackground : theme.HeaderBackground);
            canvas.DrawLine(0, PixelCentre(bounds.Bottom), rowHeaders, PixelCentre(bounds.Bottom), theme.HeaderLine);
            canvas.DrawText(row.ToString(CultureInfo.InvariantCulture), bounds, GridTextAlignment.Center, GridVerticalAlignment.Center, whole ? theme.SelectedHeaderText : theme.HeaderText, font, wrap: false);
            if (highlighted)
                underlines.Add((bounds.Y - 1, bounds.Bottom, false));
        }

        // The corner above the row headers covers whatever scrolled under it, and lines close the headers off.
        canvas.FillRectangle(new GridRect(0, 0, rowHeaders, ColumnHeaderHeight), theme.HeaderBackground);
        canvas.DrawLine(0, PixelCentre(ColumnHeaderHeight), viewport.Width, PixelCentre(ColumnHeaderHeight), theme.HeaderLine);
        canvas.DrawLine(PixelCentre(rowHeaders), 0, PixelCentre(rowHeaders), viewport.Height, theme.HeaderLine);

        // The underlines cover the two pixels before the cells, the closing line among them.
        var (underX, underY) = (Math.Floor(rowHeaders) - 1, Math.Floor(ColumnHeaderHeight) - 1);
        foreach (var (from, to, column) in underlines)
        {
            if (column)
                canvas.DrawLine(Math.Max(from, rowHeaders), underY, to, underY, theme.ActiveCellBorder, ActiveBorderThickness);
            else
                canvas.DrawLine(underX, Math.Max(from, ColumnHeaderHeight), underX, to, theme.ActiveCellBorder, ActiveBorderThickness);
        }
    }

    // The centre of the last pixel before an edge, where a one-pixel line covers exactly that pixel.
    private static double PixelCentre(double edge) => Math.Floor(edge) - 0.5;

    /// <summary>What one paint works with, worked out once: the visible cells, the merged areas among them, header width and fonts.</summary>
    private sealed class Frame
    {
        private readonly GridPainter _painter;

        public Frame(GridPainter painter, Worksheet? sheet, GridViewport viewport, double rowHeaders, Region region)
        {
            _painter = painter;
            Viewport = viewport;
            RowHeaders = rowHeaders;
            Area = region.Area;
            Clip = region.Clip;
            MinColumn = region.MinColumn;
            MaxColumn = region.MaxColumn;
            HeaderFont = DefaultFont(sheet);
            DigitWidth = painter.MaxDigitWidth(sheet);
            var merges = new List<CellArea>();
            var all = new List<CellArea>();
            if (sheet is not null)
            {
                foreach (var range in sheet.MergedAreas)
                {
                    var merge = new CellArea(range.FirstRow, range.FirstColumn, range.LastRow, range.LastColumn);
                    all.Add(merge);
                    if (Intersects(Area, merge))
                        merges.Add(merge);
                }
            }

            Merges = merges;
            _allMerges = all;
        }

        private readonly List<CellArea> _allMerges;

        public GridViewport Viewport { get; }

        public double RowHeaders { get; }

        public CellArea Area { get; }

        /// <summary>Where the region is on screen.</summary>
        public GridRect Clip { get; }

        /// <summary>The columns text may run over: those on this side of the frozen line.</summary>
        public int MinColumn { get; }

        public int MaxColumn { get; }

        public GridFont HeaderFont { get; }

        public double DigitWidth { get; }

        public IReadOnlyList<CellArea> Merges { get; }

        public CellArea? MergeAt(CellPosition cell) => MergeAt(cell.Row, cell.Column, all: false);

        /// <summary>The merged area holding a cell: among the visible ones, or among all of the sheet's.</summary>
        public CellArea? MergeAt(int row, int column, bool all)
        {
            foreach (var merge in all ? _allMerges : Merges)
            {
                if (merge.Contains(row, column))
                    return merge;
            }

            return null;
        }

        /// <summary>Whether text may run on over a cell: it holds nothing, not even a formula, and is not merged.</summary>
        public bool IsFree(Worksheet sheet, int row, int column)
        {
            var cell = sheet[row, column];
            return cell.Value.Kind == CellValueKind.Empty && cell.Formula is null && MergeAt(row, column, all: true) is null;
        }

        public GridRect Bounds(CellArea area) => _painter.Bounds(Viewport, area, RowHeaders);

        public IEnumerable<(int Row, int Column)> Cells()
        {
            for (var row = Area.FirstRow; row <= Area.LastRow; row++)
            {
                if (_painter.Layout.RowHeight(row) <= 0)
                    continue;
                for (var column = Area.FirstColumn; column <= Area.LastColumn; column++)
                {
                    if (_painter.Layout.ColumnWidth(column) > 0)
                        yield return (row, column);
                }
            }
        }
    }
}
