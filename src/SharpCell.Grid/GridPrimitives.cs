using System;

namespace SharpCell.Grid;

/// <summary>A rectangle in device-independent pixels.</summary>
/// <param name="X">The left edge.</param>
/// <param name="Y">The top edge.</param>
/// <param name="Width">The width.</param>
/// <param name="Height">The height.</param>
public readonly record struct GridRect(double X, double Y, double Width, double Height)
{
    /// <summary>The right edge.</summary>
    public double Right => X + Width;

    /// <summary>The bottom edge.</summary>
    public double Bottom => Y + Height;
}

/// <summary>An opaque or translucent sRGB colour.</summary>
/// <param name="R">Red, 0 to 255.</param>
/// <param name="G">Green, 0 to 255.</param>
/// <param name="B">Blue, 0 to 255.</param>
/// <param name="A">Opacity, 0 (transparent) to 255 (opaque).</param>
public readonly record struct GridColor(byte R, byte G, byte B, byte A = 255)
{
    /// <summary>An opaque colour from a 0xRRGGBB number.</summary>
    public static GridColor FromRgb(int rgb) => new((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb);
}

/// <summary>Where text sits across its bounds.</summary>
public enum GridTextAlignment
{
    /// <summary>Against the left edge.</summary>
    Left,

    /// <summary>In the middle.</summary>
    Center,

    /// <summary>Against the right edge.</summary>
    Right,
}

/// <summary>Where text sits from top to bottom of its bounds.</summary>
public enum GridVerticalAlignment
{
    /// <summary>Against the top edge.</summary>
    Top,

    /// <summary>In the middle.</summary>
    Center,

    /// <summary>Against the bottom edge, as Excel places cell text.</summary>
    Bottom,
}

/// <summary>
/// A cell of the sheet by 1-based row and column, such as the active cell.
/// </summary>
public readonly record struct CellPosition
{
    /// <summary>A position inside the sheet.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The row or column is outside the sheet.</exception>
    public CellPosition(int row, int column)
    {
        if (row < 1 || row > SheetLayout.MaxRow)
            throw new ArgumentOutOfRangeException(nameof(row), row, "The row is outside the sheet.");
        if (column < 1 || column > SheetLayout.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, "The column is outside the sheet.");
        Row = row;
        Column = column;
    }

    /// <summary>The row, from 1.</summary>
    public int Row { get; }

    /// <summary>The column, from 1.</summary>
    public int Column { get; }

    /// <summary>The A1 address, such as <c>B3</c>.</summary>
    public override string ToString() => ColumnNames.Of(Column) + Row.ToString(System.Globalization.CultureInfo.InvariantCulture);
}

/// <summary>
/// What part of the sheet a grid shows: the scroll offset into the sheet and the size of the whole
/// grid on screen, headers included.
/// </summary>
/// <param name="OffsetX">How far the cells are scrolled to the left, in pixels.</param>
/// <param name="OffsetY">How far the cells are scrolled up, in pixels.</param>
/// <param name="Width">The width of the grid.</param>
/// <param name="Height">The height of the grid.</param>
public readonly record struct GridViewport(double OffsetX, double OffsetY, double Width, double Height);
