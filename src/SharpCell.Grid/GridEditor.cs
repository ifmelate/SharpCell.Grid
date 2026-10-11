using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpCell.Grid;

/// <summary>
/// Changes a workbook's cells as people change them in Excel: typing into cells, Delete, pasting
/// text copied from Excel. Each change is one step to undo and redo, and the workbook is
/// recalculated after it. What is typed is read as Excel reads it, through <see cref="Cell.Input"/>.
/// Only the cells' content is kept for undoing: values, formulas and number formats, not styles.
/// </summary>
public sealed class GridEditor
{
    private readonly Stack<Step> _undo = new();
    private readonly Stack<Step> _redo = new();

    /// <summary>An editor for a workbook, with nothing to undo.</summary>
    public GridEditor(Workbook workbook)
    {
        ArgumentNullException.ThrowIfNull(workbook);
        Workbook = workbook;
    }

    /// <summary>The workbook changed.</summary>
    public Workbook Workbook { get; }

    /// <summary>Whether there is a change to undo.</summary>
    public bool CanUndo => _undo.Count > 0;

    /// <summary>Whether there is an undone change to redo.</summary>
    public bool CanRedo => _redo.Count > 0;

    /// <summary>After each change, undo and redo included, with the cells changed and what recalculating changed.</summary>
    public event EventHandler<GridEdit>? Edited;

    /// <summary>
    /// Types the same text into every cell of an area, as Enter does for one cell and Ctrl+Enter
    /// for a selection. Inside a merged area only its top-left cell takes the text.
    /// </summary>
    /// <exception cref="FormulaParseException">The text starts with <c>=</c> but is not a formula; no cell changes.</exception>
    /// <exception cref="InvalidOperationException">A cell is part of an array formula other than its top-left cell; no cell changes.</exception>
    public GridEdit Enter(Worksheet sheet, CellArea area, string input)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(input);
        var cells = new List<(int Row, int Column)>();
        var merges = sheet.MergedAreas;
        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
            {
                if (GridNavigation.MergeAt(merges, new CellPosition(row, column)) is not { } merge || (merge.FirstRow == row && merge.FirstColumn == column))
                    cells.Add((row, column));
            }
        }

        return Change(sheet, area, cells, cell => cell.Input = input);
    }

    /// <summary>
    /// Clears the content of the cells in an area, as Delete does: values and formulas go, number
    /// formats and styles stay.
    /// </summary>
    /// <exception cref="InvalidOperationException">A cell is part of an array formula other than its top-left cell; no cell changes.</exception>
    public GridEdit Clear(Worksheet sheet, CellArea area)
    {
        ArgumentNullException.ThrowIfNull(sheet);

        // Only cells with content change, so a whole column is as quick as its used part.
        var cells = sheet.Cells
            .Where(c => area.Contains(c.Row, c.Column) && (c.Formula is not null || c.Value.Kind != CellValueKind.Empty))
            .Select(c => (c.Row, c.Column))
            .ToList();
        return Change(sheet, area, cells, cell => cell.Input = "");
    }

    /// <summary>
    /// Pastes text copied from Excel or another program with its top-left cell at a cell: tabs
    /// between cells, line breaks between rows, cells with line breaks in double quotes. Each cell
    /// is read as typed; an empty one clears its cell. What lies beyond the sheet's edge is left out.
    /// </summary>
    /// <exception cref="InvalidOperationException">A cell is part of an array formula other than its top-left cell; no cell changes.</exception>
    public GridEdit Paste(Worksheet sheet, CellPosition at, string text)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(text);
        var rows = GridClipboard.ReadText(text);
        var lastRow = Math.Min(SheetLayout.MaxRow, at.Row + Math.Max(1, rows.Count) - 1);
        var lastColumn = Math.Min(SheetLayout.MaxColumn, at.Column + Math.Max(1, rows.Count == 0 ? 1 : rows.Max(r => r.Count)) - 1);
        var area = new CellArea(at.Row, at.Column, lastRow, lastColumn);
        var cells = new List<(int Row, int Column)>();
        var inputs = new Dictionary<(int, int), string>();
        for (var r = 0; r < rows.Count && at.Row + r <= lastRow; r++)
        {
            for (var c = 0; c < rows[r].Count && at.Column + c <= lastColumn; c++)
            {
                cells.Add((at.Row + r, at.Column + c));
                inputs[(at.Row + r, at.Column + c)] = rows[r][c];
            }
        }

        // A formula that does not parse is pasted as the text it is, as Excel pastes it.
        return Change(sheet, area, cells, cell =>
        {
            var input = inputs[(cell.Row, cell.Column)];
            try
            {
                cell.Input = input;
            }
            catch (FormulaParseException)
            {
                cell.Value = CellValue.Text(input);
            }
        });
    }

    /// <summary>
    /// Moves cells, as cutting and pasting them in Excel does: their values or formulas and number
    /// formats go to the area with its top-left cell at <paramref name="to"/>, and the cells left
    /// behind are cleared. Formulas move as they are written; what lies beyond the sheet's edge is
    /// left out. Styles stay where they were.
    /// </summary>
    /// <exception cref="InvalidOperationException">A cell is part of an array formula other than its top-left cell; no cell changes.</exception>
    public GridEdit Move(Worksheet sheet, CellArea from, CellPosition to)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        var (rows, columns) = (to.Row - from.FirstRow, to.Column - from.FirstColumn);
        var area = new CellArea(
            to.Row,
            to.Column,
            Math.Min(SheetLayout.MaxRow, from.LastRow + rows),
            Math.Min(SheetLayout.MaxColumn, from.LastColumn + columns));

        // Only cells with content take part: those moved, and those overwritten where the moved ones land.
        var moved = new Dictionary<(int, int), Content>();
        var cells = new HashSet<(int Row, int Column)>();
        foreach (var cell in sheet.Cells.ToList())
        {
            if (cell.Formula is null && cell.Value.Kind == CellValueKind.Empty)
                continue;
            if (from.Contains(cell.Row, cell.Column))
            {
                cells.Add((cell.Row, cell.Column));
                var target = (cell.Row + rows, cell.Column + columns);
                if (area.Contains(target.Item1, target.Item2))
                {
                    moved[target] = Content.Of(cell);
                    cells.Add(target);
                }
            }

            if (area.Contains(cell.Row, cell.Column))
                cells.Add((cell.Row, cell.Column));
        }

        var order = cells.OrderBy(c => c.Row).ThenBy(c => c.Column).ToList();
        return Change(sheet, area, order, cell =>
        {
            if (moved.TryGetValue((cell.Row, cell.Column), out var content))
            {
                content.RestoreTo(cell);
            }
            else
            {
                cell.Value = CellValue.Empty;
                cell.NumberFormat = "General";
            }
        });
    }

    /// <summary>Undoes the last change; null when there is none.</summary>
    public GridEdit? Undo() => Replay(_undo, _redo, s => s.Before);

    /// <summary>Redoes the last undone change; null when there is none.</summary>
    public GridEdit? Redo() => Replay(_redo, _undo, s => s.After);

    /// <summary>Forgets every change, so that none can be undone, as after a workbook is replaced.</summary>
    public void ClearHistory()
    {
        _undo.Clear();
        _redo.Clear();
    }

    private GridEdit Change(Worksheet sheet, CellArea area, List<(int Row, int Column)> cells, Action<Cell> change)
    {
        var before = cells.Select(c => Content.Of(sheet[c.Row, c.Column])).ToList();
        try
        {
            foreach (var (row, column) in cells)
                change(sheet[row, column]);
        }
        catch
        {
            // All or nothing: the cells changed before the failing one get their content back.
            Restore(sheet, before);
            throw;
        }

        var after = cells.Select(c => Content.Of(sheet[c.Row, c.Column])).ToList();
        _undo.Push(new Step(sheet, area, before, after));
        _redo.Clear();
        return Done(sheet, area);
    }

    private GridEdit? Replay(Stack<Step> from, Stack<Step> to, Func<Step, List<Content>> contents)
    {
        if (from.Count == 0)
            return null;
        var step = from.Pop();
        Restore(step.Sheet, contents(step));
        to.Push(step);
        return Done(step.Sheet, step.Area);
    }

    private GridEdit Done(Worksheet sheet, CellArea area)
    {
        var edit = new GridEdit(sheet, area, Workbook.Recalculate());
        Edited?.Invoke(this, edit);
        return edit;
    }

    private static void Restore(Worksheet sheet, List<Content> contents)
    {
        foreach (var content in contents)
            content.Restore(sheet);
    }

    private sealed record Step(Worksheet Sheet, CellArea Area, List<Content> Before, List<Content> After);

    // What undoing gives a cell back: its formula or value, and its number format. A cell another
    // cell's formula spills into holds nothing of its own: it is given back empty, for the spill to fill.
    private readonly record struct Content(int Row, int Column, string? Formula, CellValue Value, string NumberFormat)
    {
        public static Content Of(Cell cell)
        {
            var spilled = cell.Formula is null && cell.Value.Kind != CellValueKind.Empty && cell.Precedents.Count > 0;
            return new(cell.Row, cell.Column, cell.Formula, spilled ? CellValue.Empty : cell.Value, cell.NumberFormat);
        }

        public void Restore(Worksheet sheet) => RestoreTo(sheet[Row, Column]);

        public void RestoreTo(Cell cell)
        {
            if (Formula is not null)
                cell.Formula = Formula;
            else
                cell.Value = Value;
            cell.NumberFormat = NumberFormat;
        }
    }
}

/// <summary>A change made through a <see cref="GridEditor"/>.</summary>
/// <param name="Sheet">The sheet of the cells changed.</param>
/// <param name="Area">The cells changed: typed into, cleared or pasted over.</param>
/// <param name="Recalculation">What recalculating the workbook after the change changed, other cells included.</param>
public sealed record GridEdit(Worksheet Sheet, CellArea Area, RecalculationResult Recalculation);
