using System;

namespace SharpCell.Grid;

/// <summary>
/// Picking references with the mouse while a formula is typed, as in Excel: where the caret takes
/// one, and how a picked cell or range is written.
/// </summary>
public static class FormulaPointing
{
    // After these an operand is expected: a click there puts a reference in.
    private const string Before = "=(,+-*/^&<>:;";

    /// <summary>
    /// Whether a click on a cell puts a reference at the caret rather than entering the text: the
    /// text is a formula and the caret, past any spaces, follows <c>=</c>, an opening bracket, a
    /// comma or an operator, outside quoted text.
    /// </summary>
    /// <param name="text">The text being typed.</param>
    /// <param name="caret">Where the caret is, from 0.</param>
    public static bool CanPoint(string text, int caret)
    {
        ArgumentNullException.ThrowIfNull(text);
        if (caret < 1 || caret > text.Length || text[0] != '=')
            return false;
        var quoted = false;
        for (var i = 0; i < caret; i++)
        {
            if (text[i] == '"')
                quoted = !quoted;
        }

        if (quoted)
            return false;
        if (caret < text.Length && (char.IsLetterOrDigit(text[caret]) || text[caret] is '$' or '.' or '_'))
            return false;
        var before = caret - 1;
        while (before > 0 && text[before] == ' ')
            before--;
        return Before.Contains(text[before], StringComparison.Ordinal);
    }

    /// <summary>
    /// The reference a formula takes for cells picked on a sheet: <c>B3</c>, <c>B3:C5</c>, whole
    /// columns <c>C:D</c> or rows <c>2:2</c>, with the sheet's name in front when they are on
    /// another sheet than the formula's.
    /// </summary>
    /// <param name="area">The cells picked.</param>
    /// <param name="sheet">The sheet they are on.</param>
    /// <param name="formulaSheet">The sheet of the cell being edited.</param>
    public static string Reference(CellArea area, Worksheet? sheet, Worksheet? formulaSheet)
    {
        string reference;
        if (area.FirstRow == 1 && area.LastRow == SheetLayout.MaxRow)
            reference = $"{ColumnNames.Of(area.FirstColumn)}:{ColumnNames.Of(area.LastColumn)}";
        else if (area.FirstColumn == 1 && area.LastColumn == SheetLayout.MaxColumn)
            reference = $"{area.FirstRow}:{area.LastRow}";
        else if (area.FirstRow == area.LastRow && area.FirstColumn == area.LastColumn)
            reference = $"{ColumnNames.Of(area.FirstColumn)}{area.FirstRow}";
        else
            reference = $"{ColumnNames.Of(area.FirstColumn)}{area.FirstRow}:{ColumnNames.Of(area.LastColumn)}{area.LastRow}";
        if (sheet is null || sheet == formulaSheet)
            return reference;
        return Quoted(sheet.Name) + "!" + reference;
    }

    // A sheet name as a formula writes it: in quotes unless it is plain letters, digits and underscores
    // not starting with a digit.
    private static string Quoted(string name)
    {
        var plain = name.Length > 0 && !char.IsDigit(name[0]);
        foreach (var c in name)
            plain &= char.IsLetterOrDigit(c) || c is '_' or '.';
        return plain ? name : "'" + name.Replace("'", "''", StringComparison.Ordinal) + "'";
    }
}
