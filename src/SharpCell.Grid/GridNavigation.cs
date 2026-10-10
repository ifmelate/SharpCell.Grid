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

    /// <summary>
    /// The cell Ctrl with an arrow lands on, as in Excel: from a cell with content whose neighbour
    /// has content too, the last cell with content before an empty one; otherwise the next cell
    /// with content that way; with none, the sheet's edge. A cell has content when it holds a
    /// value or a formula.
    /// </summary>
    /// <param name="sheet">The sheet whose content is looked at; without one, the edge.</param>
    /// <param name="from">The cell to move from.</param>
    /// <param name="rows">-1 up, 1 down, or 0.</param>
    /// <param name="columns">-1 left, 1 right, or 0.</param>
    public static CellPosition MoveToEdge(Worksheet? sheet, CellPosition from, int rows, int columns)
    {
        (rows, columns) = (Math.Sign(rows), Math.Sign(columns));
        var edge = new CellPosition(
            rows < 0 ? 1 : rows > 0 ? SheetLayout.MaxRow : from.Row,
            columns < 0 ? 1 : columns > 0 ? SheetLayout.MaxColumn : from.Column);
        if (sheet is null || (rows == 0 && columns == 0))
            return sheet is null ? edge : from;

        bool Inside(int row, int column) => row is >= 1 and <= SheetLayout.MaxRow && column is >= 1 and <= SheetLayout.MaxColumn;
        bool Filled(int row, int column)
        {
            var cell = sheet[row, column];
            return cell.Formula is not null || cell.Value.Kind != CellValueKind.Empty;
        }

        var (nextRow, nextColumn) = (from.Row + rows, from.Column + columns);
        if (!Inside(nextRow, nextColumn))
            return from;

        // Along a block of content to its last cell.
        if (Filled(from.Row, from.Column) && Filled(nextRow, nextColumn))
        {
            while (Inside(nextRow + rows, nextColumn + columns) && Filled(nextRow + rows, nextColumn + columns))
                (nextRow, nextColumn) = (nextRow + rows, nextColumn + columns);
            return new CellPosition(nextRow, nextColumn);
        }

        // Over empty cells to the next with content: the nearest among the sheet's cells with content.
        CellPosition? nearest = null;
        var distance = int.MaxValue;
        foreach (var cell in sheet.Cells)
        {
            if (cell.Formula is null && cell.Value.Kind == CellValueKind.Empty)
                continue;
            var d = rows != 0
                ? (cell.Column == from.Column ? (cell.Row - from.Row) * rows : -1)
                : (cell.Row == from.Row ? (cell.Column - from.Column) * columns : -1);
            if (d > 0 && d < distance)
                (nearest, distance) = (new CellPosition(cell.Row, cell.Column), d);
        }

        return nearest ?? edge;
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
