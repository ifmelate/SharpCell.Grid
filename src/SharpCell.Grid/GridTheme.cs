namespace SharpCell.Grid;

/// <summary>The colours of a grid.</summary>
public sealed record GridTheme
{
    /// <summary>Colours close to Excel's light theme, as <see cref="Excel"/>; set properties in an initializer to change some.</summary>
    public GridTheme()
    {
    }

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

    /// <summary>Behind the headers of the rows and columns of the selected cells.</summary>
    public GridColor ActiveHeaderBackground { get; init; } = GridColor.FromRgb(0xD9D9D9);

    /// <summary>The line where panes are frozen.</summary>
    public GridColor FrozenLine { get; init; } = GridColor.FromRgb(0x9B9B9B);

    /// <summary>The border around the active cell and the selection, and the line under the headers of selected cells.</summary>
    public GridColor ActiveCellBorder { get; init; } = GridColor.FromRgb(0x107C41);

    /// <summary>The translucent shade over selected cells other than the active cell.</summary>
    public GridColor SelectionFill { get; init; } = new(0, 0, 0, 0x26);

    /// <summary>Behind the headers of whole selected columns or rows.</summary>
    public GridColor SelectedHeaderBackground { get; init; } = GridColor.FromRgb(0xC6C6C6);

    /// <summary>Text of the headers of whole selected columns or rows.</summary>
    public GridColor SelectedHeaderText { get; init; } = GridColor.FromRgb(0x107C41);

    /// <summary>The dashed border around the cells a fill handle is dragged over.</summary>
    public GridColor FillOutline { get; init; } = GridColor.FromRgb(0x7F7F7F);
}
