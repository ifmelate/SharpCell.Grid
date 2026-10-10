namespace SharpCell.Grid;

/// <summary>A rectangle of cells, from its top-left to its bottom-right cell; rows and columns from 1.</summary>
/// <param name="FirstRow">The top row.</param>
/// <param name="FirstColumn">The left column.</param>
/// <param name="LastRow">The bottom row.</param>
/// <param name="LastColumn">The right column.</param>
public readonly record struct CellArea(int FirstRow, int FirstColumn, int LastRow, int LastColumn)
{
    /// <summary>Whether the cell is inside the area.</summary>
    public bool Contains(int row, int column) =>
        row >= FirstRow && row <= LastRow && column >= FirstColumn && column <= LastColumn;
}
