namespace SharpCell.Grid;

/// <summary>The colours of a grid.</summary>
public sealed record GridTheme
{
    /// <summary>Colours close to Excel's light theme.</summary>
    public static GridTheme Excel { get; } = new();

    /// <summary>Behind the cells.</summary>
    public GridColor Background { get; init; } = GridColor.FromRgb(0xFFFFFF);

    /// <summary>The lines between cells.</summary>
    public GridColor Gridline { get; init; } = GridColor.FromRgb(0xE1E1E1);

    /// <summary>Text of cells.</summary>
    public GridColor CellText { get; init; } = GridColor.FromRgb(0x000000);

    /// <summary>Behind the row and column headers.</summary>
    public GridColor HeaderBackground { get; init; } = GridColor.FromRgb(0xF5F5F5);

    /// <summary>The lines between headers and around them.</summary>
    public GridColor HeaderLine { get; init; } = GridColor.FromRgb(0xC8C8C8);

    /// <summary>Header text: column letters and row numbers.</summary>
    public GridColor HeaderText { get; init; } = GridColor.FromRgb(0x444444);

    /// <summary>Behind the headers of the active cell's row and column.</summary>
    public GridColor ActiveHeaderBackground { get; init; } = GridColor.FromRgb(0xD9D9D9);

    /// <summary>The border around the active cell.</summary>
    public GridColor ActiveCellBorder { get; init; } = GridColor.FromRgb(0x107C41);
}
