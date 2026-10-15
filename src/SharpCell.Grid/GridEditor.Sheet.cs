using System;
using System.Collections.Generic;
using System.Linq;

namespace SharpCell.Grid;

// Changes to the sheet's rows and columns: inserting and deleting them, their sizes, hiding them.
public sealed partial class GridEditor
{
    /// <summary>
    /// Inserts empty rows before <paramref name="row"/>, as Excel's Insert Sheet Rows does: the rows
    /// below move down and formulas follow them (see <see cref="Worksheet.InsertRows"/>). The new
    /// rows take the styles and number formats of the row above, as in Excel.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The rows would not fit on the sheet.</exception>
    /// <exception cref="InvalidOperationException">Content would be pushed past the last row, or the rows would split
    /// an array formula; nothing changes.</exception>
    /// <exception cref="NotSupportedException">A formula SharpCell cannot read would have to follow; nothing changes.</exception>
    public GridEdit InsertRows(Worksheet sheet, int row, int count)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        void Insert()
        {
            sheet.InsertRows(row, count);
            if (row > 1)
                FormatLike(sheet, rows: true, row - 1, row, count);
        }

        Insert();
        return Push(new Step(sheet, new CellArea(row, 1, row + count - 1, SheetLayout.MaxColumn), () => sheet.DeleteRows(row, count), Insert, Layout: true));
    }

    /// <summary>Inserts empty columns before <paramref name="column"/>, which take the styles of the column on their left; see <see cref="InsertRows"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The columns would not fit on the sheet.</exception>
    /// <exception cref="InvalidOperationException">Content would be pushed past the last column, or the columns would
    /// split an array formula; nothing changes.</exception>
    /// <exception cref="NotSupportedException">The columns would go inside a table, or a formula SharpCell cannot read
    /// would have to follow; nothing changes.</exception>
    public GridEdit InsertColumns(Worksheet sheet, int column, int count)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        void Insert()
        {
            sheet.InsertColumns(column, count);
            if (column > 1)
                FormatLike(sheet, rows: false, column - 1, column, count);
        }

        Insert();
        return Push(new Step(sheet, new CellArea(1, column, SheetLayout.MaxRow, column + count - 1), () => sheet.DeleteColumns(column, count), Insert, Layout: true));
    }

    /// <summary>
    /// Deletes rows, as Excel's Delete Sheet Rows does: the rows below move up, and references to
    /// the deleted cells become <c>#REF!</c> (see <see cref="Worksheet.DeleteRows"/>). Undoing gives
    /// back the rows with their cells, styles and heights, and every formula and name as it was.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The rows are not all on the sheet.</exception>
    /// <exception cref="InvalidOperationException">The rows hold part of an array formula, or part of a table that
    /// cannot go; nothing changes.</exception>
    /// <exception cref="NotSupportedException">A formula SharpCell cannot read would have to follow; nothing changes.</exception>
    public GridEdit DeleteRows(Worksheet sheet, int row, int count)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        return Delete(sheet, rows: true, row, count);
    }

    /// <summary>Deletes columns; see <see cref="DeleteRows"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The columns are not all on the sheet.</exception>
    /// <exception cref="InvalidOperationException">The columns hold part of an array formula; nothing changes.</exception>
    /// <exception cref="NotSupportedException">The columns are part of a table, or a formula SharpCell cannot read
    /// would have to follow; nothing changes.</exception>
    public GridEdit DeleteColumns(Worksheet sheet, int column, int count)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        return Delete(sheet, rows: false, column, count);
    }

    /// <summary>
    /// Inserts empty cells in place of an area, as Excel's Insert with "Shift cells down" or "Shift
    /// cells right" does: only the cells beside the area move, and formulas follow them (see
    /// <see cref="CellRange.Insert"/>). The new cells take the styles and number formats of the
    /// cells above (or on the left), as in Excel. An area of whole rows or columns inserts them.
    /// </summary>
    /// <exception cref="InvalidOperationException">Content would be pushed off the sheet, or the cells that move would
    /// split a merged area, a table or an array formula; nothing changes.</exception>
    /// <exception cref="NotSupportedException">A formula SharpCell cannot read would have to follow; nothing changes.</exception>
    public GridEdit InsertCells(Worksheet sheet, CellArea area, InsertShift shift)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        var down = shift == InsertShift.Down;
        if (down && IsWholeRows(area))
            return InsertRows(sheet, area.FirstRow, area.LastRow - area.FirstRow + 1);
        if (!down && IsWholeColumns(area))
            return InsertColumns(sheet, area.FirstColumn, area.LastColumn - area.FirstColumn + 1);
        var range = RangeOf(sheet, area);
        void Insert()
        {
            range.Insert(shift);
            FormatCellsLike(sheet, area, down);
        }

        Insert();
        return Push(new Step(sheet, area, () => range.Delete(down ? DeleteShift.Up : DeleteShift.Left), Insert, Layout: true));
    }

    /// <summary>
    /// Deletes the cells of an area, as Excel's Delete with "Shift cells up" or "Shift cells left"
    /// does: the cells below (or on the right) move into their place, and references to the deleted
    /// cells become <c>#REF!</c> (see <see cref="CellRange.Delete"/>). Undoing gives back the cells and
    /// every formula and name as it was. An area of whole rows or columns deletes them.
    /// </summary>
    /// <exception cref="InvalidOperationException">The cells that move would split a merged area, a table or an array
    /// formula; nothing changes.</exception>
    /// <exception cref="NotSupportedException">A formula SharpCell cannot read would have to follow; nothing changes.</exception>
    public GridEdit DeleteCells(Worksheet sheet, CellArea area, DeleteShift shift)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        var up = shift == DeleteShift.Up;
        if (up && IsWholeRows(area))
            return DeleteRows(sheet, area.FirstRow, area.LastRow - area.FirstRow + 1);
        if (!up && IsWholeColumns(area))
            return DeleteColumns(sheet, area.FirstColumn, area.LastColumn - area.FirstColumn + 1);
        var range = RangeOf(sheet, area);
        void Remove() => range.Delete(shift);
        var kept = Deleted.Keep(sheet, area, () => range.Insert(up ? InsertShift.Down : InsertShift.Right));
        Remove();
        if (kept is null)
        {
            ClearHistory();
            return Done(new Step(sheet, area, () => { }, () => { }, Layout: true));
        }

        return Push(new Step(sheet, area, kept.Restore, Remove, Layout: true));
    }

    private static bool IsWholeRows(CellArea area) => area.FirstColumn == 1 && area.LastColumn == SheetLayout.MaxColumn;

    private static bool IsWholeColumns(CellArea area) => area.FirstRow == 1 && area.LastRow == SheetLayout.MaxRow;

    // Inserted cells look like their neighbours above, or on the left: their styles and number formats.
    private static void FormatCellsLike(Worksheet sheet, CellArea area, bool down)
    {
        if (down ? area.FirstRow == 1 : area.FirstColumn == 1)
            return;
        var defaultStyle = sheet.Workbook.DefaultStyle;
        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
            {
                var from = down ? sheet[area.FirstRow - 1, column] : sheet[row, area.FirstColumn - 1];
                var target = sheet[row, column];
                if (!ReferenceEquals(from.Style, defaultStyle))
                    target.Style = from.Style;
                if (from.NumberFormat != "General")
                    target.NumberFormat = from.NumberFormat;
            }
        }
    }

    /// <summary>
    /// Gives columns a width, as dragging a column header's edge does in Excel: in the units of
    /// <see cref="Worksheet.SetColumnWidth"/>, or null for the sheet's default.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">A column is outside the sheet, or the width is outside 0 to 255.</exception>
    public GridEdit SetColumnWidth(Worksheet sheet, int firstColumn, int lastColumn, double? width)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        return SetColumnWidths(sheet, firstColumn, lastColumn, _ => width);
    }

    /// <summary>
    /// Fits columns to their content, as double-clicking a column header's edge does in Excel: each
    /// takes the width the painter finds for its widest text (see <see cref="GridPainter.FitColumnWidth"/>);
    /// a column with nothing to fit keeps its width.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">A column is outside the sheet.</exception>
    public GridEdit FitColumnWidth(Worksheet sheet, int firstColumn, int lastColumn, GridPainter painter)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(painter);
        Check(firstColumn, lastColumn, SheetLayout.MaxColumn, nameof(lastColumn));
        var digit = painter.MaxDigitWidth(sheet);
        var widths = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1)
            .Select(c => painter.FitColumnWidth(sheet, c) is { } pixels ? Math.Min(255, SheetLayout.ColumnWidthFromPixels(pixels, digit)) : sheet.ColumnWidth(c))
            .ToArray();
        return SetColumnWidths(sheet, firstColumn, lastColumn, c => widths[c - firstColumn]);
    }

    private GridEdit SetColumnWidths(Worksheet sheet, int firstColumn, int lastColumn, Func<int, double?> width)
    {
        Check(firstColumn, lastColumn, SheetLayout.MaxColumn, nameof(lastColumn));
        var before = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1).Select(sheet.ColumnWidth).ToArray();
        void Apply(Func<int, double?> size)
        {
            for (var column = firstColumn; column <= lastColumn; column++)
                sheet.SetColumnWidth(column, size(column));
        }

        try
        {
            Apply(width);
        }
        catch (ArgumentOutOfRangeException)
        {
            Apply(c => before[c - firstColumn]);
            throw;
        }

        return Push(new Step(sheet, new CellArea(1, firstColumn, SheetLayout.MaxRow, lastColumn),
            () => Apply(c => before[c - firstColumn]), () => Apply(width), Layout: true));
    }

    /// <summary>Gives rows a height in points, or null for the sheet's default; see <see cref="SetColumnWidth"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A row is outside the sheet, or the height is outside 0 to 409.</exception>
    public GridEdit SetRowHeight(Worksheet sheet, int firstRow, int lastRow, double? height)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        return SetRowHeights(sheet, firstRow, lastRow, _ => height);
    }

    /// <summary>
    /// Fits rows to their content, as double-clicking a row header's edge does in Excel: each takes
    /// the height the painter finds for its tallest text (see <see cref="GridPainter.FitRowHeight"/>),
    /// and a row with nothing in it the sheet's default height.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">A row is outside the sheet.</exception>
    public GridEdit FitRowHeight(Worksheet sheet, int firstRow, int lastRow, GridPainter painter)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(painter);
        Check(firstRow, lastRow, SheetLayout.MaxRow, nameof(lastRow));
        var heights = Enumerable.Range(firstRow, lastRow - firstRow + 1)
            .Select(r => painter.FitRowHeight(sheet, r) is { } pixels ? SheetLayout.RowHeightFromPixels(pixels, painter.Zoom) : (double?)null)
            .ToArray();
        return SetRowHeights(sheet, firstRow, lastRow, r => heights[r - firstRow]);
    }

    private GridEdit SetRowHeights(Worksheet sheet, int firstRow, int lastRow, Func<int, double?> height)
    {
        Check(firstRow, lastRow, SheetLayout.MaxRow, nameof(lastRow));
        var before = Enumerable.Range(firstRow, lastRow - firstRow + 1).Select(sheet.RowHeight).ToArray();
        void Apply(Func<int, double?> size)
        {
            for (var row = firstRow; row <= lastRow; row++)
                sheet.SetRowHeight(row, size(row));
        }

        try
        {
            Apply(height);
        }
        catch (ArgumentOutOfRangeException)
        {
            Apply(r => before[r - firstRow]);
            throw;
        }

        return Push(new Step(sheet, new CellArea(firstRow, 1, lastRow, SheetLayout.MaxColumn),
            () => Apply(r => before[r - firstRow]), () => Apply(height), Layout: true));
    }

    /// <summary>Hides or shows columns, as Excel's Hide and Unhide do; a hidden column keeps its width for when it shows again.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A column is outside the sheet.</exception>
    public GridEdit SetColumnHidden(Worksheet sheet, int firstColumn, int lastColumn, bool hidden)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        Check(firstColumn, lastColumn, SheetLayout.MaxColumn, nameof(lastColumn));
        var before = Enumerable.Range(firstColumn, lastColumn - firstColumn + 1).Select(sheet.IsColumnHidden).ToArray();
        void Apply(Func<int, bool> hide)
        {
            for (var column = firstColumn; column <= lastColumn; column++)
                sheet.SetColumnHidden(column, hide(column));
        }

        Apply(_ => hidden);
        return Push(new Step(sheet, new CellArea(1, firstColumn, SheetLayout.MaxRow, lastColumn),
            () => Apply(c => before[c - firstColumn]), () => Apply(_ => hidden), Layout: true));
    }

    /// <summary>Hides or shows rows; see <see cref="SetColumnHidden"/>.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A row is outside the sheet.</exception>
    public GridEdit SetRowHidden(Worksheet sheet, int firstRow, int lastRow, bool hidden)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        Check(firstRow, lastRow, SheetLayout.MaxRow, nameof(lastRow));
        var before = Enumerable.Range(firstRow, lastRow - firstRow + 1).Select(sheet.IsRowHidden).ToArray();
        void Apply(Func<int, bool> hide)
        {
            for (var row = firstRow; row <= lastRow; row++)
                sheet.SetRowHidden(row, hide(row));
        }

        Apply(_ => hidden);
        return Push(new Step(sheet, new CellArea(firstRow, 1, lastRow, SheetLayout.MaxColumn),
            () => Apply(r => before[r - firstRow]), () => Apply(_ => hidden), Layout: true));
    }

    private static void Check(int first, int last, int max, string parameter)
    {
        if (first < 1 || first > max)
            throw new ArgumentOutOfRangeException(parameter == "lastRow" ? "firstRow" : "firstColumn", first, "The position is outside the sheet.");
        if (last < first || last > max)
            throw new ArgumentOutOfRangeException(parameter, last, "The last position must be on the sheet, at or after the first.");
    }

    private GridEdit Delete(Worksheet sheet, bool rows, int at, int count)
    {
        var area = rows ? new CellArea(at, 1, at + count - 1, SheetLayout.MaxColumn) : new CellArea(1, at, SheetLayout.MaxRow, at + count - 1);
        void Remove()
        {
            if (rows)
                sheet.DeleteRows(at, count);
            else
                sheet.DeleteColumns(at, count);
        }

        // What the deleted cells held is kept aside first; a sheet whose cells cannot be kept that
        // way, as with array formulas, is changed without a way back.
        var kept = Deleted.Keep(sheet, area, () =>
        {
            if (rows)
                sheet.InsertRows(at, count);
            else
                sheet.InsertColumns(at, count);
        }, (rows, at, count));
        Remove();
        if (kept is null)
        {
            ClearHistory();
            return Done(new Step(sheet, area, () => { }, () => { }, Layout: true));
        }

        return Push(new Step(sheet, area, kept.Restore, Remove, Layout: true));
    }

    // Inserted rows or columns look like their neighbour: its cells' styles and number formats.
    private static void FormatLike(Worksheet sheet, bool rows, int from, int at, int count)
    {
        var length = rows ? SheetLayout.MaxColumn : SheetLayout.MaxRow;
        var defaultStyle = sheet.Workbook.DefaultStyle;
        for (var i = 1; i <= length; i++)
        {
            var cell = rows ? sheet[from, i] : sheet[i, from];
            var style = cell.Style;
            var format = cell.NumberFormat;
            var styled = !ReferenceEquals(style, defaultStyle);
            var formatted = format != "General";
            if (!styled && !formatted)
                continue;
            for (var k = 0; k < count; k++)
            {
                var target = rows ? sheet[at + k, i] : sheet[i, at + k];
                if (styled)
                    target.Style = style;
                if (formatted)
                    target.NumberFormat = format;
            }
        }
    }

    // Everything deleting rows, columns or cells changes, kept to give it back: the deleted cells,
    // the sizes of deleted rows or columns, every formula, name and data validation rule of the
    // workbook as written before, merged areas and tables.
    private sealed class Deleted
    {
        private readonly Worksheet _sheet;
        private readonly Action _reinsert;
        private readonly (bool Rows, int At, int Count)? _lines;
        private readonly CellRange _cells;
        private readonly CellArea _area;
        private readonly double?[] _sizes;
        private readonly bool[] _hidden;
        private readonly List<(Worksheet Sheet, int Row, int Column, string Formula)> _formulas = [];
        private readonly List<DefinedName> _names;
        private readonly List<string> _merges;
        private readonly List<(string Name, string Range, bool Header, bool Totals)> _tables;
        private readonly List<(Worksheet Sheet, Rules Rules)> _rules;

        private Deleted(Worksheet sheet, CellArea area, CellRange cells, Action reinsert, (bool Rows, int At, int Count)? lines)
        {
            (_sheet, _area, _cells, _reinsert, _lines) = (sheet, area, cells, reinsert, lines);
            var (rows, at, count) = lines ?? (false, 0, 0);
            _sizes = new double?[count];
            _hidden = new bool[count];
            for (var i = 0; i < count; i++)
            {
                _sizes[i] = rows ? sheet.RowHeight(at + i) : sheet.ColumnWidth(at + i);
                _hidden[i] = rows ? sheet.IsRowHidden(at + i) : sheet.IsColumnHidden(at + i);
            }

            foreach (var other in sheet.Workbook.Sheets)
            {
                foreach (var cell in other.Cells)
                {
                    if (cell.Formula is { } formula && !(other == sheet && area.Contains(cell.Row, cell.Column)))
                        _formulas.Add((other, cell.Row, cell.Column, formula));
                }
            }

            _names = [.. sheet.Workbook.DefinedNames];
            _merges = sheet.MergedAreas.Select(m => m.Address).ToList();
            _tables = sheet.Workbook.Tables.Where(t => t.Worksheet == sheet).Select(t => (t.Name, t.Range, t.HasHeaderRow, t.HasTotalsRow)).ToList();
            _rules = sheet.Workbook.Sheets.Select(s => (s, Rules.Of(s))).ToList();
        }

        /// <param name="sheet">The sheet.</param>
        /// <param name="area">The cells deleted.</param>
        /// <param name="reinsert">Inserts empty cells where they were.</param>
        /// <param name="lines">For whole rows or columns, which; their sizes are kept too.</param>
        public static Deleted? Keep(Worksheet sheet, CellArea area, Action reinsert, (bool Rows, int At, int Count)? lines = null)
        {
            // The deleted cells are copied into a workbook of their own, at the same addresses.
            var copy = new Workbook().AddSheet("Deleted");
            try
            {
                var band = RangeOf(sheet, area);
                band.CopyTo(RangeOf(copy, area));
                return new Deleted(sheet, area, RangeOf(copy, area), reinsert, lines);
            }
            catch (NotSupportedException)
            {
                return null;
            }
            catch (ArgumentOutOfRangeException)
            {
                // Rows or columns not on the sheet: deleting them fails as it should.
                return null;
            }
        }

        public void Restore()
        {
            _reinsert();
            _cells.CopyTo(RangeOf(_sheet, _area));
            var (rows, at, _) = _lines ?? (false, 0, 0);
            for (var i = 0; i < _sizes.Length; i++)
            {
                if (rows)
                {
                    _sheet.SetRowHeight(at + i, _sizes[i]);
                    _sheet.SetRowHidden(at + i, _hidden[i]);
                }
                else
                {
                    _sheet.SetColumnWidth(at + i, _sizes[i]);
                    _sheet.SetColumnHidden(at + i, _hidden[i]);
                }
            }

            // Formulas and names that lost references, or whose ranges grew back differently, are written again.
            foreach (var (sheet, row, column, formula) in _formulas)
            {
                var cell = sheet[row, column];
                if (cell.Formula != formula)
                    cell.Formula = formula;
            }

            foreach (var name in _names)
            {
                if (!_sheet.Workbook.DefinedNames.Any(n => n.Name == name.Name && n.Scope == name.Scope && n.Formula == name.Formula))
                    _sheet.Workbook.DefineName(name.Name, name.Formula, name.Scope);
            }

            var merges = _sheet.MergedAreas.Select(m => m.Address).ToList();
            foreach (var merge in merges.Except(_merges))
                _sheet.Unmerge(merge);
            foreach (var merge in _merges.Except(merges))
                _sheet.Merge(merge);

            foreach (var (name, range, header, totals) in _tables)
            {
                if (!_sheet.Workbook.TryGetTable(name, out _))
                    _sheet.AddTable(name, range, header, totals);
            }

            // Data validation as it was, on this sheet and on sheets whose rules read it.
            foreach (var (sheet, rules) in _rules)
                rules.Restore(sheet);
        }
    }
}
