using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace SharpCell.Grid;

/// <summary>Which edges of an area's cells a border goes on, as in Excel's Borders menu.</summary>
[Flags]
public enum BorderEdges
{
    /// <summary>No edge.</summary>
    None = 0,

    /// <summary>The left edge of the area.</summary>
    Left = 1,

    /// <summary>The right edge of the area.</summary>
    Right = 2,

    /// <summary>The top edge of the area.</summary>
    Top = 4,

    /// <summary>The bottom edge of the area.</summary>
    Bottom = 8,

    /// <summary>The edges between the area's columns.</summary>
    InsideVertical = 16,

    /// <summary>The edges between the area's rows.</summary>
    InsideHorizontal = 32,

    /// <summary>Round the area: Excel's Outside Borders.</summary>
    Outside = Left | Right | Top | Bottom,

    /// <summary>Between the cells inside the area.</summary>
    Inside = InsideVertical | InsideHorizontal,

    /// <summary>Every edge of every cell: Excel's All Borders.</summary>
    All = Outside | Inside,
}

// Changing how cells look, as Excel's Home tab does: fonts, fills, borders and number formats.
public sealed partial class GridEditor
{
    /// <summary>
    /// Changes the style of every cell of an area, as a button of Excel's Home tab does: the
    /// function gets each cell's style and gives the new one, such as <c>style =&gt; style with { Fill = color }</c>.
    /// Whole rows or columns change their cells up to the sheet's used range.
    /// </summary>
    /// <param name="sheet">The sheet.</param>
    /// <param name="area">The cells.</param>
    /// <param name="change">The new style of a cell from its old one.</param>
    public GridEdit SetStyle(Worksheet sheet, CellArea area, Func<CellStyle, CellStyle> change)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(change);
        var cells = Formatted(sheet, area);
        return Format(sheet, area, cells, () =>
        {
            foreach (var (row, column) in cells)
            {
                var cell = sheet[row, column];
                var style = change(cell.Style);
                if (!Equals(style, cell.Style))
                    cell.Style = style;
            }
        });
    }

    /// <summary>
    /// Gives every cell of an area a number format, as Excel's Number Format box does; see
    /// <see cref="Cell.NumberFormat"/>. Whole rows or columns change their cells up to the sheet's used range.
    /// </summary>
    /// <exception cref="ArgumentException">The format is not a number format Excel reads.</exception>
    public GridEdit SetNumberFormat(Worksheet sheet, CellArea area, string format)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(format);
        var cells = Formatted(sheet, area);
        return Format(sheet, area, cells, () =>
        {
            foreach (var (row, column) in cells)
                sheet[row, column].NumberFormat = format;
        });
    }

    /// <summary>
    /// Puts a border on edges of an area, or takes it off with null, as Excel's Borders menu does:
    /// <see cref="BorderEdges.Outside"/> for Outside Borders, <see cref="BorderEdges.All"/> for All
    /// Borders, <see cref="BorderEdges.All"/> with null for No Border. The neighbouring cell's edge
    /// on the other side of a line goes, so the line looks as chosen.
    /// </summary>
    /// <param name="sheet">The sheet.</param>
    /// <param name="area">The cells.</param>
    /// <param name="edges">Which edges.</param>
    /// <param name="border">The line, such as <c>new CellBorder(CellBorderStyle.Thin)</c>; null for none.</param>
    public GridEdit SetBorders(Worksheet sheet, CellArea area, BorderEdges edges, CellBorder? border)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        var used = Formatted(sheet, area);
        if (used.Count == 0)
            return Format(sheet, area, used, () => { });
        var box = new CellArea(used.Min(c => c.Row), used.Min(c => c.Column), used.Max(c => c.Row), used.Max(c => c.Column));
        var set = new Dictionary<(int Row, int Column), Func<CellStyle, CellStyle>>();
        void Edge((int, int) cell, Func<CellStyle, CellStyle> change)
        {
            if (cell.Item1 < 1 || cell.Item2 < 1 || cell.Item1 > SheetLayout.MaxRow || cell.Item2 > SheetLayout.MaxColumn)
                return;
            set[cell] = set.TryGetValue(cell, out var before) ? style => change(before(style)) : change;
        }

        foreach (var (row, column) in used)
        {
            if (edges.HasFlag(row == box.FirstRow ? BorderEdges.Top : BorderEdges.InsideHorizontal))
            {
                Edge((row, column), s => s with { TopBorder = border });
                if (row == box.FirstRow)
                    Edge((row - 1, column), s => s with { BottomBorder = null });
            }

            if (edges.HasFlag(row == box.LastRow ? BorderEdges.Bottom : BorderEdges.InsideHorizontal))
            {
                Edge((row, column), s => s with { BottomBorder = border });
                if (row == box.LastRow)
                    Edge((row + 1, column), s => s with { TopBorder = null });
            }

            if (edges.HasFlag(column == box.FirstColumn ? BorderEdges.Left : BorderEdges.InsideVertical))
            {
                Edge((row, column), s => s with { LeftBorder = border });
                if (column == box.FirstColumn)
                    Edge((row, column - 1), s => s with { RightBorder = null });
            }

            if (edges.HasFlag(column == box.LastColumn ? BorderEdges.Right : BorderEdges.InsideVertical))
            {
                Edge((row, column), s => s with { RightBorder = border });
                if (column == box.LastColumn)
                    Edge((row, column + 1), s => s with { LeftBorder = null });
            }
        }

        // The neighbours outside change only where they had a facing edge.
        var cells = set.Keys.Where(c => box.Contains(c.Row, c.Column) || HasEdge(sheet[c.Row, c.Column].Style)).ToList();
        return Format(sheet, area, cells, () =>
        {
            foreach (var cell in cells)
            {
                var target = sheet[cell.Row, cell.Column];
                var style = set[cell](target.Style);
                if (!Equals(style, target.Style))
                    target.Style = style;
            }
        });

        static bool HasEdge(CellStyle style) => style.LeftBorder is not null || style.RightBorder is not null || style.TopBorder is not null || style.BottomBorder is not null;
    }

    /// <summary>
    /// The number formats of Excel's Ctrl+Shift shortcuts, for a workbook's culture: <c>~</c> General,
    /// <c>!</c> two decimals with thousands separators, <c>$</c> currency, <c>%</c> percent, <c>^</c>
    /// scientific, <c>#</c> a date with day, month and year, <c>@</c> a time with hours and minutes.
    /// </summary>
    /// <param name="symbol">The shortcut's character.</param>
    /// <param name="culture">The workbook's culture, for the currency symbol.</param>
    /// <returns>The format code; null for another character.</returns>
    public static string? ShortcutFormat(char symbol, CultureInfo culture)
    {
        ArgumentNullException.ThrowIfNull(culture);
        return symbol switch
        {
            '~' => "General",
            '!' => "#,##0.00",
            '$' => culture.NumberFormat.CurrencySymbol == "$" ? "$#,##0.00_);($#,##0.00)" : "#,##0.00 \"" + culture.NumberFormat.CurrencySymbol + "\"",
            '%' => "0%",
            '^' => "0.00E+00",
            '#' => "d-mmm-yy",
            '@' => "h:mm AM/PM",
            _ => null,
        };
    }

    // The cells a format change touches: every cell of the area, or for whole rows and columns
    // those inside the sheet's used range.
    private static List<(int Row, int Column)> Formatted(Worksheet sheet, CellArea area)
    {
        if (area.FirstRow == 1 && area.LastRow == SheetLayout.MaxRow || area.FirstColumn == 1 && area.LastColumn == SheetLayout.MaxColumn)
        {
            if (sheet.UsedRange is not { } used)
                return [];
            area = new CellArea(Math.Max(area.FirstRow, used.FirstRow), Math.Max(area.FirstColumn, used.FirstColumn),
                Math.Min(area.LastRow, used.LastRow), Math.Min(area.LastColumn, used.LastColumn));
            if (area.FirstRow > area.LastRow || area.FirstColumn > area.LastColumn)
                return [];
        }

        if ((long)(area.LastRow - area.FirstRow + 1) * (area.LastColumn - area.FirstColumn + 1) > MaxRepeatedCells)
            throw new InvalidOperationException("The area has too many cells to format at once.");
        return Every(area);
    }

    // A change of styles or number formats, undone by giving the cells back theirs.
    private GridEdit Format(Worksheet sheet, CellArea area, List<(int Row, int Column)> cells, Action change)
    {
        var defaultStyle = sheet.Workbook.DefaultStyle;
        (CellStyle? Style, string Format)[] Look() => [.. cells.Select(c =>
        {
            var cell = sheet[c.Row, c.Column];
            return (ReferenceEquals(cell.Style, defaultStyle) ? null : cell.Style, cell.NumberFormat);
        })];
        void Apply((CellStyle? Style, string Format)[] look)
        {
            for (var i = 0; i < cells.Count; i++)
            {
                var cell = sheet[cells[i].Row, cells[i].Column];
                if (!ReferenceEquals(cell.Style, look[i].Style ?? defaultStyle))
                    cell.Style = look[i].Style;
                if (cell.NumberFormat != look[i].Format)
                    cell.NumberFormat = look[i].Format;
            }
        }

        var before = Look();
        try
        {
            change();
        }
        catch
        {
            Apply(before);
            throw;
        }

        var after = Look();
        return Push(new Step(sheet, area, () => Apply(before), () => Apply(after), Layout: false));
    }
}
