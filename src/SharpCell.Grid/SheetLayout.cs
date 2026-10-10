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
