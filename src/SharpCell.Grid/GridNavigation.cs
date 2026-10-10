using System;
using System.Collections.Generic;

namespace SharpCell.Grid;

/// <summary>How the active cell moves on a sheet, as in Excel: a merged area counts as one cell.</summary>
public static class GridNavigation
{
    /// <summary>
    /// The cell a move by whole rows and columns lands on, kept inside the sheet. Leaving a merged
    /// area starts from its far edge; landing inside one gives its top-left cell.
    /// </summary>
    public static CellPosition Move(Worksheet? sheet, CellPosition from, int rows, int columns)
    {
        var merges = sheet?.MergedAreas ?? [];
        var start = MergeAt(merges, from);
        var row = from.Row + rows;
        var column = from.Column + columns;
        if (start is { } area)
        {
            if (rows > 0)
                row = area.LastRow + rows;
            if (columns > 0)
                column = area.LastColumn + columns;
            if (rows < 0)
                row = area.FirstRow + rows;
            if (columns < 0)
                column = area.FirstColumn + columns;
        }

        var target = new CellPosition(Math.Clamp(row, 1, SheetLayout.MaxRow), Math.Clamp(column, 1, SheetLayout.MaxColumn));
        return MergeAt(merges, target) is { } landed ? new CellPosition(landed.FirstRow, landed.FirstColumn) : target;
    }

    /// <summary>The merged area holding a cell, if any.</summary>
    internal static CellArea? MergeAt(IReadOnlyList<CellRange> merges, CellPosition cell)
    {
        foreach (var range in merges)
        {
            if (cell.Row >= range.FirstRow && cell.Row <= range.LastRow && cell.Column >= range.FirstColumn && cell.Column <= range.LastColumn)
                return new CellArea(range.FirstRow, range.FirstColumn, range.LastRow, range.LastColumn);
        }

        return null;
    }
}
