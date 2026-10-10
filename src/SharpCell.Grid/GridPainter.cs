using System;
using System.Globalization;

namespace SharpCell.Grid;

/// <summary>
/// Paints a worksheet the way Excel shows it: column and row headers, gridlines, each cell's
/// <see cref="Cell.Text"/> with Excel's General alignment, and the active cell. Only the cells
/// inside the viewport are read, so a sheet's size does not matter. Also answers the geometry
/// questions of a grid on screen: where a cell is, what is under a point, where to scroll.
/// With <see cref="MeasureText"/> set, numbers too wide for their column are shortened or shown as
/// <c>###</c>, and row headers widen for long row numbers, as in Excel.
/// </summary>
public sealed class GridPainter
{
    // Space between a cell's edges and its text.
    private const double TextPadding = 2;
    private const double ActiveBorderThickness = 2;

    // Space on each side of a row number, and the digits row headers always have room for.
    private const double HeaderPadding = 6;
    private const int MinRowHeaderDigits = 3;

    /// <summary>A painter with Excel's default row heights and column widths.</summary>
    public GridPainter()
        : this(new SheetLayout())
    {
    }

    /// <summary>A painter over the given sizes of rows and columns.</summary>
    public GridPainter(SheetLayout layout)
    {
        ArgumentNullException.ThrowIfNull(layout);
        Layout = layout;
    }

    /// <summary>The sizes of rows and columns.</summary>
    public SheetLayout Layout { get; }

    /// <summary>The colours.</summary>
    public GridTheme Theme { get; set; } = GridTheme.Excel;

    /// <summary>The height of the column headers at the top.</summary>
    public double ColumnHeaderHeight { get; set; } = 20;

    /// <summary>
    /// The width of the row headers on the left; with <see cref="MeasureText"/>, the least width,
    /// widened to fit the number of the last visible row.
    /// </summary>
    public double RowHeaderWidth { get; set; } = 40;

    /// <summary>
    /// The width of a line of text in the font cells are painted with, in pixels. Without it, text
    /// is never fitted to its cell and row headers keep <see cref="RowHeaderWidth"/>.
    /// </summary>
    public Func<string, double>? MeasureText { get; set; }

    /// <summary>Paints the part of the sheet in the viewport; without a sheet, an empty grid.</summary>
    public void Paint(IGridCanvas canvas, Worksheet? sheet, GridViewport viewport, CellPosition? activeCell = null)
    {
        ArgumentNullException.ThrowIfNull(canvas);
        var theme = Theme;
        canvas.FillRectangle(new GridRect(0, 0, viewport.Width, viewport.Height), theme.Background);
        if (CellsWidth(viewport) <= 0 || CellsHeight(viewport) <= 0)
            return;

        var area = VisibleArea(viewport);
        if (sheet is not null)
            PaintCells(canvas, sheet, viewport, area);
        PaintGridlines(canvas, viewport, area);
        PaintHeaders(canvas, viewport, area, activeCell);
        if (activeCell is { } active && area.Contains(active.Row, active.Column))
        {
            var bounds = CellBounds(viewport, active);
            canvas.StrokeRectangle(
                new GridRect(bounds.X - 0.5, bounds.Y - 0.5, bounds.Width, bounds.Height),
                theme.ActiveCellBorder,
                ActiveBorderThickness);
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
    public GridRect CellBounds(GridViewport viewport, CellPosition cell) => CellBounds(viewport, cell, RowHeaders(viewport));

    private GridRect CellBounds(GridViewport viewport, CellPosition cell, double rowHeaders) => new(
        rowHeaders + Layout.ColumnLeft(cell.Column) - viewport.OffsetX,
        ColumnHeaderHeight + Layout.RowTop(cell.Row) - viewport.OffsetY,
        Layout.ColumnWidth(cell.Column),
        Layout.RowHeight(cell.Row));

    /// <summary>The cell under a point of the grid; null on the headers or outside the grid.</summary>
    public CellPosition? HitTest(GridViewport viewport, double x, double y)
    {
        if (x < RowHeaders(viewport) || y < ColumnHeaderHeight || x >= viewport.Width || y >= viewport.Height)
            return null;
        return new CellPosition(
            Layout.RowAt(y - ColumnHeaderHeight + viewport.OffsetY),
            Layout.ColumnAt(x - RowHeaders(viewport) + viewport.OffsetX));
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

    private double CellsWidth(GridViewport viewport) => viewport.Width - RowHeaders(viewport);

    // Excel widens row headers for the digits of the visible row numbers. Rows do not depend on
    // the headers' width, so the last visible row is found first.
    private double RowHeaders(GridViewport viewport)
    {
        if (MeasureText is not { } measure)
            return RowHeaderWidth;
        var height = Math.Max(0, CellsHeight(viewport));
        var lastRow = Layout.RowAt(Math.Max(0, viewport.OffsetY) + (height == 0 ? 0 : Math.BitDecrement(height)));
        var digits = Math.Max(MinRowHeaderDigits, lastRow.ToString(CultureInfo.InvariantCulture).Length);
        return Math.Max(RowHeaderWidth, Math.Ceiling(measure(new string('0', digits))) + 2 * HeaderPadding);
    }

    private double CellsHeight(GridViewport viewport) => viewport.Height - ColumnHeaderHeight;

    private void PaintCells(IGridCanvas canvas, Worksheet sheet, GridViewport viewport, CellArea area)
    {
        var rowHeaders = RowHeaders(viewport);
        var color = Theme.CellText;
        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
            {
                var cell = sheet[row, column];
                var text = cell.Text;
                if (text.Length == 0)
                    continue;
                var bounds = CellBounds(viewport, new CellPosition(row, column), rowHeaders);
                if (bounds.Width <= 0 || bounds.Height <= 0)
                    continue;

                // The text keeps clear of the gridline on the cell's right and bottom edges.
                var inner = new GridRect(
                    bounds.X + TextPadding,
                    bounds.Y,
                    Math.Max(0, bounds.Width - 2 * TextPadding - 1),
                    Math.Max(0, bounds.Height - 1));
                var value = cell.Value;
                if (value.Kind == CellValueKind.Number && MeasureText is { } measure)
                {
                    var general = cell.NumberFormat.Equals("General", StringComparison.OrdinalIgnoreCase);
                    text = TextFitting.FitNumber(text, value.AsNumber(), general, sheet.Workbook.Culture, inner.Width, measure);
                }
                canvas.DrawText(text, inner, AlignmentOf(value.Kind), GridVerticalAlignment.Bottom, color);
            }
        }
    }

    // Excel's General alignment: numbers (dates too) right, text left, logical values and errors centred.
    private static GridTextAlignment AlignmentOf(CellValueKind kind) => kind switch
    {
        CellValueKind.Number => GridTextAlignment.Right,
        CellValueKind.Boolean or CellValueKind.Error => GridTextAlignment.Center,
        _ => GridTextAlignment.Left,
    };

    private void PaintGridlines(IGridCanvas canvas, GridViewport viewport, CellArea area)
    {
        var rowHeaders = RowHeaders(viewport);
        var color = Theme.Gridline;
        for (var column = area.FirstColumn; column <= area.LastColumn; column++)
        {
            var right = rowHeaders + Layout.ColumnLeft(column) + Layout.ColumnWidth(column) - viewport.OffsetX;
            if (Layout.ColumnWidth(column) > 0)
                canvas.DrawLine(PixelCentre(right), ColumnHeaderHeight, PixelCentre(right), viewport.Height, color);
        }

        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            var bottom = ColumnHeaderHeight + Layout.RowTop(row) + Layout.RowHeight(row) - viewport.OffsetY;
            if (Layout.RowHeight(row) > 0)
                canvas.DrawLine(rowHeaders, PixelCentre(bottom), viewport.Width, PixelCentre(bottom), color);
        }
    }

    private void PaintHeaders(IGridCanvas canvas, GridViewport viewport, CellArea area, CellPosition? activeCell)
    {
        var rowHeaders = RowHeaders(viewport);
        var theme = Theme;
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
            canvas.DrawText(ColumnNames.Of(column), bounds, GridTextAlignment.Center, GridVerticalAlignment.Center, theme.HeaderText);
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
            canvas.DrawText(row.ToString(CultureInfo.InvariantCulture), bounds, GridTextAlignment.Center, GridVerticalAlignment.Center, theme.HeaderText);
        }

        // The corner above the row headers covers whatever scrolled under it, and lines close the headers off.
        canvas.FillRectangle(new GridRect(0, 0, rowHeaders, ColumnHeaderHeight), theme.HeaderBackground);
        canvas.DrawLine(0, PixelCentre(ColumnHeaderHeight), viewport.Width, PixelCentre(ColumnHeaderHeight), theme.HeaderLine);
        canvas.DrawLine(PixelCentre(rowHeaders), 0, PixelCentre(rowHeaders), viewport.Height, theme.HeaderLine);
    }

    // The centre of the last pixel before an edge, where a one-pixel line covers exactly that pixel.
    private static double PixelCentre(double edge) => Math.Floor(edge) - 0.5;
}
