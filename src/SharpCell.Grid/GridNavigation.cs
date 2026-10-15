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
    /// The current region of a cell, as Excel's Ctrl+A and Ctrl+* select it: the cell grown, row by
    /// row and column by column, while a cell with content touches it, also diagonally. A cell with
    /// nothing around it is its own region.
    /// </summary>
    /// <param name="sheet">The sheet whose content is looked at; without one, the cell.</param>
    /// <param name="cell">The cell to start from.</param>
    public static CellArea CurrentRegion(Worksheet? sheet, CellPosition cell)
    {
        var area = new CellArea(cell.Row, cell.Column, cell.Row, cell.Column);
        if (sheet is null)
            return area;
        while (true)
        {
            var ring = new CellArea(Math.Max(1, area.FirstRow - 1), Math.Max(1, area.FirstColumn - 1),
                Math.Min(SheetLayout.MaxRow, area.LastRow + 1), Math.Min(SheetLayout.MaxColumn, area.LastColumn + 1));
            var grown = area;
            foreach (var strip in Strips(ring, area))
            {
                foreach (var found in sheet.Range(strip).Cells)
                {
                    if (found.Formula is null && found.Value.Kind == CellValueKind.Empty)
                        continue;
                    grown = new CellArea(Math.Min(grown.FirstRow, found.Row), Math.Min(grown.FirstColumn, found.Column),
                        Math.Max(grown.LastRow, found.Row), Math.Max(grown.LastColumn, found.Column));
                }
            }

            if (grown == area)
                return area;
            area = grown;
        }

        // The ring around an area as up to four ranges: the rows above and below, the columns beside.
        static IEnumerable<string> Strips(CellArea ring, CellArea area)
        {
            if (ring.FirstRow < area.FirstRow)
                yield return Address(new CellArea(ring.FirstRow, ring.FirstColumn, ring.FirstRow, ring.LastColumn));
            if (ring.LastRow > area.LastRow)
                yield return Address(new CellArea(ring.LastRow, ring.FirstColumn, ring.LastRow, ring.LastColumn));
            if (ring.FirstColumn < area.FirstColumn)
                yield return Address(new CellArea(area.FirstRow, ring.FirstColumn, area.LastRow, ring.FirstColumn));
            if (ring.LastColumn > area.LastColumn)
                yield return Address(new CellArea(area.FirstRow, ring.LastColumn, area.LastRow, ring.LastColumn));
        }

        static string Address(CellArea area) =>
            $"{ColumnNames.Of(area.FirstColumn)}{area.FirstRow}:{ColumnNames.Of(area.LastColumn)}{area.LastRow}";
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
