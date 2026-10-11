using System;

namespace SharpCell.Grid;

/// <summary>
/// Where the cells of a sheet are, in device-independent pixels from the top-left corner of cell
/// A1: column widths, row heights, their offsets, the cell under a point and the cells inside a
/// viewport. Covers Excel's whole sheet, 1,048,576 rows by 16,384 columns. Rows and columns are
/// numbered from 1.
/// </summary>
public sealed class SheetLayout
{
    /// <summary>Excel's default column width at 100% zoom: 8.43 characters of Calibri 11, 64 pixels.</summary>
    public const double ExcelColumnWidth = 64;

    /// <summary>Excel's default row height at 100% zoom: 15 points of Calibri 11, 20 pixels.</summary>
    public const double ExcelRowHeight = 20;

    /// <summary>Rows of an Excel sheet.</summary>
    public const int MaxRow = 1_048_576;

    /// <summary>Columns of an Excel sheet, A to XFD.</summary>
    public const int MaxColumn = 16_384;

    private readonly AxisLayout columns;
    private readonly AxisLayout rows;

    /// <summary>A layout with Excel's default sizes.</summary>
    public SheetLayout()
        : this(ExcelColumnWidth, ExcelRowHeight)
    {
    }

    /// <summary>A layout with the given default sizes, in pixels.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A size is not positive.</exception>
    public SheetLayout(double defaultColumnWidth, double defaultRowHeight)
    {
        columns = new AxisLayout(MaxColumn, defaultColumnWidth, "column");
        rows = new AxisLayout(MaxRow, defaultRowHeight, "row");
    }

    /// <summary>
    /// A layout with a worksheet's column widths, row heights and hidden rows and columns, in
    /// pixels at 100% zoom as Excel converts them: a width of <c>w</c> characters is
    /// <c>Truncate((256 * w + Truncate(128 / d)) / 256 * d)</c> pixels, <c>d</c> being the width of
    /// the widest digit of the workbook's default font (7 for Calibri 11); a height of <c>h</c>
    /// points is <c>h * 96 / 72</c> pixels, rounded. Later changes to the sheet are not followed.
    /// </summary>
    /// <param name="sheet">The sheet.</param>
    /// <param name="maxDigitWidth">The width in pixels of the widest digit of the default font.</param>
    /// <exception cref="ArgumentOutOfRangeException">The digit width is not positive.</exception>
    public static SheetLayout FromWorksheet(Worksheet sheet, double maxDigitWidth) => FromWorksheet(sheet, maxDigitWidth, 1);

    /// <summary>
    /// A layout with a worksheet's sizes at a zoom, as Excel shows them: column widths from the
    /// digit width of the default font at that zoom, row heights in points times the zoom.
    /// </summary>
    /// <param name="sheet">The sheet.</param>
    /// <param name="maxDigitWidth">The width in pixels of the widest digit of the default font at the zoom.</param>
    /// <param name="zoom">The zoom, 1 for 100%.</param>
    /// <exception cref="ArgumentOutOfRangeException">The digit width or the zoom is not positive.</exception>
    public static SheetLayout FromWorksheet(Worksheet sheet, double maxDigitWidth, double zoom)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        if (!(maxDigitWidth > 0) || double.IsInfinity(maxDigitWidth))
            throw new ArgumentOutOfRangeException(nameof(maxDigitWidth), maxDigitWidth, "The digit width must be positive.");
        if (!(zoom > 0) || double.IsInfinity(zoom))
            throw new ArgumentOutOfRangeException(nameof(zoom), zoom, "The zoom must be positive.");

        // Without a default width Excel uses 8 digits and 5 pixels of padding, rounded up to a multiple of 8.
        var columnWidth = sheet.DefaultColumnWidth is { } width
            ? ColumnPixels(width, maxDigitWidth)
            : Math.Ceiling((8 * maxDigitWidth + 5) / 8) * 8;
        var rowHeight = sheet.DefaultRowHeight is { } height ? RowPixels(height, zoom) : Math.Round(ExcelRowHeight * zoom);
        var layout = new SheetLayout(Math.Max(1, columnWidth), Math.Max(1, rowHeight));

        foreach (var column in Each(MaxColumn, sheet.ColumnWidth))
            layout.SetColumnWidth(column.Index, ColumnPixels(column.Size, maxDigitWidth));
        foreach (var row in Each(MaxRow, sheet.RowHeight))
            layout.SetRowHeight(row.Index, RowPixels(row.Size, zoom));
        return layout.HideFrom(sheet);
    }

    /// <summary>
    /// The column width, in the units of <see cref="Worksheet.SetColumnWidth"/>, that Excel gives a
    /// column dragged to a width in pixels: <c>Truncate((c * d + 5) / d * 256) / 256</c> for
    /// <c>c = Truncate((pixels - 5) / d * 100 + 0.5) / 100</c> characters, as Excel shows it while
    /// dragging. Narrower than a digit and its padding, the smallest width that shows the pixels.
    /// </summary>
    /// <param name="pixels">The width in pixels.</param>
    /// <param name="maxDigitWidth">The width in pixels of the widest digit of the default font, at the zoom the pixels are at.</param>
    /// <exception cref="ArgumentOutOfRangeException">The digit width is not positive, or the pixels are negative.</exception>
    public static double ColumnWidthFromPixels(double pixels, double maxDigitWidth)
    {
        if (!(maxDigitWidth > 0) || double.IsInfinity(maxDigitWidth))
            throw new ArgumentOutOfRangeException(nameof(maxDigitWidth), maxDigitWidth, "The digit width must be positive.");
        if (!(pixels >= 0))
            throw new ArgumentOutOfRangeException(nameof(pixels), pixels, "The width must be zero or more.");

        pixels = Math.Round(pixels);
        if (pixels >= maxDigitWidth + 5)
        {
            var characters = Math.Truncate((pixels - 5) / maxDigitWidth * 100 + 0.5) / 100;
            return Math.Min(255, Math.Truncate((characters * maxDigitWidth + 5) / maxDigitWidth * 256) / 256);
        }

        // The pixels a width shows only grow with it, so the steps are searched.
        int low = 0, high = 255 * 256;
        while (low < high)
        {
            var middle = (low + high) / 2;
            if (ColumnPixels(middle / 256.0, maxDigitWidth) >= pixels)
                high = middle;
            else
                low = middle + 1;
        }

        return low / 256.0;
    }

    /// <summary>The row height in points, as Excel stores it, of a row shown that many pixels high at a zoom.</summary>
    /// <param name="pixels">The height in pixels.</param>
    /// <param name="zoom">The zoom the pixels are at, 1 for 100%.</param>
    /// <exception cref="ArgumentOutOfRangeException">The zoom is not positive, or the pixels are negative.</exception>
    public static double RowHeightFromPixels(double pixels, double zoom)
    {
        if (!(zoom > 0) || double.IsInfinity(zoom))
            throw new ArgumentOutOfRangeException(nameof(zoom), zoom, "The zoom must be positive.");
        if (!(pixels >= 0))
            throw new ArgumentOutOfRangeException(nameof(pixels), pixels, "The height must be zero or more.");
        return Math.Min(409, Math.Round(pixels / zoom) * 72 / 96);
    }

    private static double ColumnPixels(double width, double digit) =>
        width <= 0 ? 0 : Math.Truncate((256 * width + Math.Truncate(128 / digit)) / 256 * digit);

    private static double RowPixels(double points, double zoom) => Math.Round(points * 96 / 72 * zoom, MidpointRounding.AwayFromZero);

    // The worksheet offers no enumeration of its sizes, so every row and column is asked: a few
    // milliseconds for a whole sheet, paid once per layout.
    private static System.Collections.Generic.IEnumerable<(int Index, double Size)> Each(int count, Func<int, double?> size)
    {
        for (var i = 1; i <= count; i++)
        {
            if (size(i) is { } value)
                yield return (i, value);
        }
    }

    private SheetLayout HideFrom(Worksheet sheet)
    {
        for (var column = 1; column <= MaxColumn; column++)
        {
            if (sheet.IsColumnHidden(column))
                SetColumnWidth(column, 0);
        }

        for (var row = 1; row <= MaxRow; row++)
        {
            if (sheet.IsRowHidden(row))
                SetRowHeight(row, 0);
        }

        return this;
    }

    /// <summary>The width of every column without a width of its own.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is not positive.</exception>
    public double DefaultColumnWidth
    {
        get => columns.DefaultSize;
        set => columns.DefaultSize = value;
    }

    /// <summary>The height of every row without a height of its own.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The height is not positive.</exception>
    public double DefaultRowHeight
    {
        get => rows.DefaultSize;
        set => rows.DefaultSize = value;
    }

    /// <summary>The width of all columns together.</summary>
    public double Width => columns.Total;

    /// <summary>The height of all rows together.</summary>
    public double Height => rows.Total;

    /// <summary>The width of a column.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    public double ColumnWidth(int column) => columns.SizeOf(column);

    /// <summary>Gives a column a width; zero hides it, the default width forgets the change.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet, or the width is negative.</exception>
    public void SetColumnWidth(int column, double width) => columns.SetSize(column, width);

    /// <summary>The height of a row.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    public double RowHeight(int row) => rows.SizeOf(row);

    /// <summary>Gives a row a height; zero hides it, the default height forgets the change.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet, or the height is negative.</exception>
    public void SetRowHeight(int row, double height) => rows.SetSize(row, height);

    /// <summary>The offset of a column's left edge; 0 for column A.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The column is outside the sheet.</exception>
    public double ColumnLeft(int column) => columns.OffsetOf(column);

    /// <summary>The offset of a row's top edge; 0 for row 1.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row is outside the sheet.</exception>
    public double RowTop(int row) => rows.OffsetOf(row);

    /// <summary>The column under a horizontal offset, clamped to the sheet; hidden columns are never hit.</summary>
    public int ColumnAt(double x) => columns.IndexAt(x);

    /// <summary>The row under a vertical offset, clamped to the sheet; hidden rows are never hit.</summary>
    public int RowAt(double y) => rows.IndexAt(y);

    /// <summary>The cells that a rectangle of the sheet touches, partly or wholly.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The width or height is negative.</exception>
    public CellArea VisibleArea(double x, double y, double width, double height)
    {
        if (!(width >= 0))
            throw new ArgumentOutOfRangeException(nameof(width), width, "The width must be zero or more.");
        if (!(height >= 0))
            throw new ArgumentOutOfRangeException(nameof(height), height, "The height must be zero or more.");

        // The far edge itself belongs to the next cell, which is not visible.
        var lastColumn = width == 0 ? ColumnAt(x) : ColumnAt(Math.BitDecrement(x + width));
        var lastRow = height == 0 ? RowAt(y) : RowAt(Math.BitDecrement(y + height));
        return new CellArea(RowAt(y), ColumnAt(x), lastRow, lastColumn);
    }
}
