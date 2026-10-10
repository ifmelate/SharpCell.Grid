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

    /// <summary>Paints the part of the sheet in the viewport; without a sheet, an empty grid.</summary>
    public void Paint(IGridCanvas canvas, Worksheet? sheet, GridViewport viewport, CellPosition? activeCell = null)
    {
        ArgumentNullException.ThrowIfNull(canvas);
        var theme = Theme;
        canvas.FillRectangle(new GridRect(0, 0, viewport.Width, viewport.Height), theme.Background);
        if (CellsWidth(viewport) <= 0 || CellsHeight(viewport) <= 0)
            return;

        var frame = new Frame(this, sheet, viewport);
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

        PaintHeaders(canvas, frame, activeCell);
        if (activeCell is { } active)
        {
            var area = frame.MergeAt(active) ?? new CellArea(active.Row, active.Column, active.Row, active.Column);
            if (Intersects(frame.Area, area))
            {
                var bounds = frame.Bounds(area);
                canvas.StrokeRectangle(
                    new GridRect(bounds.X - 0.5, bounds.Y - 0.5, bounds.Width, bounds.Height),
                    theme.ActiveCellBorder,
                    ActiveBorderThickness);
            }
        }
    }

    /// <summary>The cells the viewport shows, partly or wholly.</summary>
    public CellArea VisibleArea(GridViewport viewport) =>
        Layout.VisibleArea(
            Math.Max(0, viewport.OffsetX),
            Math.Max(0, viewport.OffsetY),
            Math.Max(0, CellsWidth(viewport)),
            Math.Max(0, CellsHeight(viewport)));

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
        var cell = new CellPosition(
            Layout.RowAt(y - ColumnHeaderHeight + viewport.OffsetY),
            Layout.ColumnAt(x - rowHeaders + viewport.OffsetX));
        return sheet is not null && GridNavigation.MergeAt(sheet.MergedAreas, cell) is { } merge
            ? new CellPosition(merge.FirstRow, merge.FirstColumn)
            : cell;
    }

    /// <summary>
    /// The viewport scrolled just enough to show the whole cell, or its top-left part when it is
    /// larger than the viewport; unchanged when the cell is already in view.
    /// </summary>
    public GridViewport ScrollIntoView(GridViewport viewport, CellPosition cell)
    {
        var offsetX = Scroll(viewport.OffsetX, CellsWidth(viewport), Layout.ColumnLeft(cell.Column), Layout.ColumnWidth(cell.Column));
        var offsetY = Scroll(viewport.OffsetY, CellsHeight(viewport), Layout.RowTop(cell.Row), Layout.RowHeight(cell.Row));
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
        var left = Layout.ColumnLeft(area.FirstColumn);
        var top = Layout.RowTop(area.FirstRow);
        return new GridRect(
            rowHeaders + left - viewport.OffsetX,
            ColumnHeaderHeight + top - viewport.OffsetY,
            Layout.ColumnLeft(area.LastColumn) + Layout.ColumnWidth(area.LastColumn) - left,
            Layout.RowTop(area.LastRow) + Layout.RowHeight(area.LastRow) - top);
    }

    // Excel widens row headers for the digits of the visible row numbers. Rows do not depend on
    // the headers' width, so the last visible row is found first.
    private double RowHeaders(GridViewport viewport)
    {
        if (MeasureText is not { } measure)
            return RowHeaderWidth;
        var height = Math.Max(0, CellsHeight(viewport));
        var lastRow = Layout.RowAt(Math.Max(0, viewport.OffsetY) + (height == 0 ? 0 : Math.BitDecrement(height)));
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
        for (var column = area.FirstColumn; column <= area.LastColumn; column++)
        {
            if (Layout.ColumnWidth(column) <= 0)
                continue;
            var right = frame.RowHeaders + Layout.ColumnLeft(column) + Layout.ColumnWidth(column) - viewport.OffsetX;
            canvas.DrawLine(PixelCentre(right), ColumnHeaderHeight, PixelCentre(right), viewport.Height, color);
        }

        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            if (Layout.RowHeight(row) <= 0)
                continue;
            var bottom = ColumnHeaderHeight + Layout.RowTop(row) + Layout.RowHeight(row) - viewport.OffsetY;
            canvas.DrawLine(frame.RowHeaders, PixelCentre(bottom), viewport.Width, PixelCentre(bottom), color);
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
        for (var column = start; column >= 1 && column <= SheetLayout.MaxColumn && Math.Abs(column - start) < OverflowReach; column += step)
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
            if (next < 1 || next > SheetLayout.MaxColumn || !frame.IsFree(sheet, row, next))
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

    private void PaintHeaders(IGridCanvas canvas, Frame frame, CellPosition? activeCell)
    {
        var theme = Theme;
        var viewport = frame.Viewport;
        var area = frame.Area;
        var rowHeaders = frame.RowHeaders;
        var font = frame.HeaderFont;
        canvas.FillRectangle(new GridRect(0, 0, viewport.Width, ColumnHeaderHeight), theme.HeaderBackground);
        canvas.FillRectangle(new GridRect(0, 0, rowHeaders, viewport.Height), theme.HeaderBackground);

        for (var column = area.FirstColumn; column <= area.LastColumn; column++)
        {
            var width = Layout.ColumnWidth(column);
            if (width <= 0)
                continue;
            var bounds = new GridRect(rowHeaders + Layout.ColumnLeft(column) - viewport.OffsetX, 0, width, ColumnHeaderHeight);
            if (activeCell?.Column == column)
                canvas.FillRectangle(bounds, theme.ActiveHeaderBackground);
            canvas.DrawLine(PixelCentre(bounds.Right), 0, PixelCentre(bounds.Right), ColumnHeaderHeight, theme.HeaderLine);
            canvas.DrawText(ColumnNames.Of(column), bounds, GridTextAlignment.Center, GridVerticalAlignment.Center, theme.HeaderText, font, wrap: false);
        }

        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            var height = Layout.RowHeight(row);
            if (height <= 0)
                continue;
            var bounds = new GridRect(0, ColumnHeaderHeight + Layout.RowTop(row) - viewport.OffsetY, rowHeaders, height);
            if (activeCell?.Row == row)
                canvas.FillRectangle(bounds, theme.ActiveHeaderBackground);
            canvas.DrawLine(0, PixelCentre(bounds.Bottom), rowHeaders, PixelCentre(bounds.Bottom), theme.HeaderLine);
            canvas.DrawText(row.ToString(CultureInfo.InvariantCulture), bounds, GridTextAlignment.Center, GridVerticalAlignment.Center, theme.HeaderText, font, wrap: false);
        }

        // The corner above the row headers covers whatever scrolled under it, and lines close the headers off.
        canvas.FillRectangle(new GridRect(0, 0, rowHeaders, ColumnHeaderHeight), theme.HeaderBackground);
        canvas.DrawLine(0, PixelCentre(ColumnHeaderHeight), viewport.Width, PixelCentre(ColumnHeaderHeight), theme.HeaderLine);
        canvas.DrawLine(PixelCentre(rowHeaders), 0, PixelCentre(rowHeaders), viewport.Height, theme.HeaderLine);
    }

    // The centre of the last pixel before an edge, where a one-pixel line covers exactly that pixel.
    private static double PixelCentre(double edge) => Math.Floor(edge) - 0.5;

    /// <summary>What one paint works with, worked out once: the visible cells, the merged areas among them, header width and fonts.</summary>
    private sealed class Frame
    {
        private readonly GridPainter _painter;

        public Frame(GridPainter painter, Worksheet? sheet, GridViewport viewport)
        {
            _painter = painter;
            Viewport = viewport;
            RowHeaders = painter.RowHeaders(viewport);
            Area = painter.VisibleArea(viewport);
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
