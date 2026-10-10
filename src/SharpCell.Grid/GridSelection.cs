using System;
using System.Collections.Generic;
using System.Globalization;

namespace SharpCell.Grid;

/// <summary>
/// The selected cells, as in Excel: a rectangle of cells and the active cell inside it. The
/// rectangle runs from the <see cref="Anchor"/>, where the selection started, to the
/// <see cref="Extent"/>, the corner the mouse or Shift with the arrows moves, and grows to cover
/// whole every merged area it touches. Whole columns span every row of the sheet, whole rows every
/// column. A selection never changes; each step gives a new one.
/// </summary>
public sealed record GridSelection
{
    private GridSelection(CellPosition activeCell, CellPosition anchor, CellPosition extent, Worksheet? sheet)
    {
        ActiveCell = activeCell;
        Anchor = anchor;
        Extent = extent;
        var merges = sheet?.MergedAreas ?? [];
        Area = Cover(Box(anchor, extent), merges);
        IsSingleCell = Area == Cover(Box(activeCell, activeCell), merges);
    }

    /// <summary>The cell typing would go into, with the white background inside the shaded selection.</summary>
    public CellPosition ActiveCell { get; }

    /// <summary>The corner the selection started from; it stays when the selection is extended.</summary>
    public CellPosition Anchor { get; }

    /// <summary>The corner opposite the anchor, which extending the selection moves.</summary>
    public CellPosition Extent { get; }

    /// <summary>The selected cells, merged areas they touch included.</summary>
    public CellArea Area { get; }

    /// <summary>Whether only the active cell is selected; a merged area counts as one cell.</summary>
    public bool IsSingleCell { get; }

    /// <summary>Whether the selection spans every row: whole columns, or the whole sheet.</summary>
    public bool IsEntireColumns => Area.FirstRow == 1 && Area.LastRow == SheetLayout.MaxRow;

    /// <summary>Whether the selection spans every column: whole rows, or the whole sheet.</summary>
    public bool IsEntireRows => Area.FirstColumn == 1 && Area.LastColumn == SheetLayout.MaxColumn;

    /// <summary>One cell, or the merged area holding it.</summary>
    public static GridSelection Cell(CellPosition cell, Worksheet? sheet) => new(cell, cell, cell, sheet);

    /// <summary>Whole columns, from the first to the last, with the active cell in the first column.</summary>
    /// <param name="first">The column the selection starts from.</param>
    /// <param name="last">The column it extends to; before <paramref name="first"/> for a selection extended to the left.</param>
    /// <param name="activeRow">The row of the active cell: in Excel, the top row on screen.</param>
    /// <param name="sheet">The sheet, whose merged areas the selection covers whole.</param>
    public static GridSelection Columns(int first, int last, int activeRow, Worksheet? sheet) =>
        new(new CellPosition(activeRow, first), new CellPosition(1, first), new CellPosition(SheetLayout.MaxRow, last), sheet);

    /// <summary>Whole rows, from the first to the last, with the active cell in the first row.</summary>
    /// <param name="first">The row the selection starts from.</param>
    /// <param name="last">The row it extends to; above <paramref name="first"/> for a selection extended upwards.</param>
    /// <param name="activeColumn">The column of the active cell: in Excel, the leftmost column on screen.</param>
    /// <param name="sheet">The sheet, whose merged areas the selection covers whole.</param>
    public static GridSelection Rows(int first, int last, int activeColumn, Worksheet? sheet) =>
        new(new CellPosition(first, activeColumn), new CellPosition(first, 1), new CellPosition(last, SheetLayout.MaxColumn), sheet);

    /// <summary>The whole sheet, keeping the active cell where it is.</summary>
    public static GridSelection All(CellPosition activeCell) =>
        new(activeCell, new CellPosition(1, 1), new CellPosition(SheetLayout.MaxRow, SheetLayout.MaxColumn), null);

    /// <summary>
    /// What a click selects: a cell, whole columns from a column header, whole rows from a row
    /// header, the whole sheet from the corner above the row headers.
    /// </summary>
    public static GridSelection At(GridHit hit, Worksheet? sheet) => hit.Part switch
    {
        GridPart.ColumnHeader => Columns(hit.Cell.Column, hit.Cell.Column, hit.Cell.Row, sheet),
        GridPart.RowHeader => Rows(hit.Cell.Row, hit.Cell.Row, hit.Cell.Column, sheet),
        GridPart.Corner => All(hit.Cell),
        _ => Cell(hit.Cell, sheet),
    };

    /// <summary>
    /// The selection from the same anchor to a cell, as when dragging the mouse or Shift-clicking:
    /// whole columns or rows stay whole, the active cell stays.
    /// </summary>
    public GridSelection ExtendTo(CellPosition cell, Worksheet? sheet)
    {
        var row = IsEntireColumns ? Extent.Row : cell.Row;
        var column = IsEntireRows ? Extent.Column : cell.Column;
        return new GridSelection(ActiveCell, Anchor, new CellPosition(row, column), sheet);
    }

    /// <summary>
    /// The selection from the same anchor to what a Shift-click hits: to a cell, to whole columns
    /// from a column header, to whole rows from a row header, the whole sheet from the corner.
    /// </summary>
    public GridSelection ExtendTo(GridHit hit, Worksheet? sheet) => hit.Part switch
    {
        GridPart.ColumnHeader => new GridSelection(ActiveCell, new CellPosition(1, Anchor.Column), new CellPosition(SheetLayout.MaxRow, hit.Cell.Column), sheet),
        GridPart.RowHeader => new GridSelection(ActiveCell, new CellPosition(Anchor.Row, 1), new CellPosition(hit.Cell.Row, SheetLayout.MaxColumn), sheet),
        GridPart.Corner => All(ActiveCell),
        _ => ExtendTo(hit.Cell, sheet),
    };

    /// <summary>
    /// The selection grown or shrunk by whole rows and columns at its extent, as Shift with the
    /// arrows does; a step that a merged area swallows goes on until the selection changes.
    /// </summary>
    public GridSelection Extend(int rows, int columns, Worksheet? sheet)
    {
        var extent = Extent;
        while (true)
        {
            var next = new CellPosition(
                Math.Clamp(extent.Row + rows, 1, SheetLayout.MaxRow),
                Math.Clamp(extent.Column + columns, 1, SheetLayout.MaxColumn));
            if (next == extent)
                return this;
            var extended = new GridSelection(ActiveCell, Anchor, next, sheet);
            if (extended.Area != Area)
                return extended;
            extent = next;
        }
    }

    /// <summary>The active cell moved by whole rows and columns, as the arrows move it, and selected alone.</summary>
    public GridSelection Move(int rows, int columns, Worksheet? sheet) =>
        Cell(GridNavigation.Move(sheet, ActiveCell, rows, columns), sheet);

    /// <summary>The active cell moved to the edge of the content, as Ctrl with an arrow moves it, and selected alone.</summary>
    /// <param name="rows">-1 up, 1 down, or 0.</param>
    /// <param name="columns">-1 left, 1 right, or 0.</param>
    /// <param name="sheet">The sheet whose content decides the edge; see <see cref="GridNavigation.MoveToEdge"/>.</param>
    public GridSelection MoveToEdge(int rows, int columns, Worksheet? sheet) =>
        Cell(GridNavigation.MoveToEdge(sheet, ActiveCell, rows, columns), sheet);

    /// <summary>The selection extended from its far corner to the edge of the content, as Ctrl+Shift with an arrow does.</summary>
    /// <param name="rows">-1 up, 1 down, or 0.</param>
    /// <param name="columns">-1 left, 1 right, or 0.</param>
    /// <param name="sheet">The sheet whose content decides the edge; see <see cref="GridNavigation.MoveToEdge"/>.</param>
    public GridSelection ExtendToEdge(int rows, int columns, Worksheet? sheet) =>
        ExtendTo(GridNavigation.MoveToEdge(sheet, Extent, rows, columns), sheet);

    /// <summary>
    /// The active cell moved as Enter and Tab move it. A single cell moves as with the arrows; inside
    /// a larger selection it moves to the next cell of the selection, which stays, and wraps around:
    /// Enter goes down each column in turn, Tab across each row. Hidden cells and the inner cells of
    /// merged areas are passed over.
    /// </summary>
    /// <param name="rows">1 for Enter, -1 for Shift+Enter, 0 for Tab.</param>
    /// <param name="columns">1 for Tab, -1 for Shift+Tab, 0 for Enter.</param>
    /// <param name="sheet">The sheet, whose hidden rows and columns and merged areas are passed over.</param>
    public GridSelection MoveWithin(int rows, int columns, Worksheet? sheet)
    {
        if (IsSingleCell || (rows == 0 && columns == 0))
            return Move(rows, columns, sheet);

        var area = Area;
        var merges = sheet?.MergedAreas ?? [];
        var byColumn = rows != 0;
        var forward = (byColumn ? rows : columns) > 0;
        var (row, column) = (ActiveCell.Row, ActiveCell.Column);
        var cells = (long)(area.LastRow - area.FirstRow + 1) * (area.LastColumn - area.FirstColumn + 1);
        for (long step = 0; step < cells; step++)
        {
            (row, column) = byColumn ? Next(row, column, area.FirstRow, area.LastRow, area.FirstColumn, area.LastColumn, forward) : Swap(Next(column, row, area.FirstColumn, area.LastColumn, area.FirstRow, area.LastRow, forward));
            if (sheet is not null && (sheet.IsRowHidden(row) || sheet.IsColumnHidden(column)))
                continue;
            var cell = new CellPosition(row, column);
            if (GridNavigation.MergeAt(merges, cell) is { } merge && (merge.FirstRow != row || merge.FirstColumn != column))
                continue;
            return new GridSelection(cell, Anchor, Extent, sheet);
        }

        return this;
    }

    /// <summary>Whole columns under the selection, keeping the active cell, as Ctrl+Space does.</summary>
    public GridSelection EntireColumns(Worksheet? sheet) =>
        new(ActiveCell, new CellPosition(1, Anchor.Column), new CellPosition(SheetLayout.MaxRow, Extent.Column), sheet);

    /// <summary>Whole rows under the selection, keeping the active cell, as Shift+Space does.</summary>
    public GridSelection EntireRows(Worksheet? sheet) =>
        new(ActiveCell, new CellPosition(Anchor.Row, 1), new CellPosition(Extent.Row, SheetLayout.MaxColumn), sheet);

    /// <summary>Whether a cell is selected.</summary>
    public bool Contains(int row, int column) => Area.Contains(row, column);

    /// <summary>The selected cells' address: <c>B3</c>, <c>B3:D7</c>, <c>C:E</c> for whole columns, <c>2:4</c> for whole rows.</summary>
    public override string ToString()
    {
        var area = Area;
        if (IsEntireRows)
            return Rows(area.FirstRow) + ":" + Rows(area.LastRow);
        if (IsEntireColumns)
            return ColumnNames.Of(area.FirstColumn) + ":" + ColumnNames.Of(area.LastColumn);
        var first = new CellPosition(area.FirstRow, area.FirstColumn).ToString();
        return area.FirstRow == area.LastRow && area.FirstColumn == area.LastColumn
            ? first
            : first + ":" + new CellPosition(area.LastRow, area.LastColumn);

        static string Rows(int row) => row.ToString(CultureInfo.InvariantCulture);
    }

    // The next cell along a line of cells, onto the next line at the end, back to the start after
    // the last line.
    private static (int Line, int Across) Next(int line, int across, int firstLine, int lastLine, int firstAcross, int lastAcross, bool forward)
    {
        if (forward)
        {
            if (line < lastLine)
                return (line + 1, across);
            return (firstLine, across < lastAcross ? across + 1 : firstAcross);
        }

        if (line > firstLine)
            return (line - 1, across);
        return (lastLine, across > firstAcross ? across - 1 : lastAcross);
    }

    private static (int Row, int Column) Swap((int Column, int Row) cell) => (cell.Row, cell.Column);

    internal static CellArea Box(CellPosition a, CellPosition b) =>
        new(Math.Min(a.Row, b.Row), Math.Min(a.Column, b.Column), Math.Max(a.Row, b.Row), Math.Max(a.Column, b.Column));

    // Grows the area until every merged area it touches lies inside it.
    internal static CellArea Cover(CellArea area, IReadOnlyList<CellRange> merges)
    {
        bool grown;
        do
        {
            grown = false;
            foreach (var merge in merges)
            {
                var touches = merge.FirstRow <= area.LastRow && area.FirstRow <= merge.LastRow && merge.FirstColumn <= area.LastColumn && area.FirstColumn <= merge.LastColumn;
                if (!touches || area.Contains(merge.FirstRow, merge.FirstColumn) && area.Contains(merge.LastRow, merge.LastColumn))
                    continue;
                area = new CellArea(
                    Math.Min(area.FirstRow, merge.FirstRow),
                    Math.Min(area.FirstColumn, merge.FirstColumn),
                    Math.Max(area.LastRow, merge.LastRow),
                    Math.Max(area.LastColumn, merge.LastColumn));
                grown = true;
            }
        }
        while (grown);
        return area;
    }
}
