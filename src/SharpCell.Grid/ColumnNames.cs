using System;

namespace SharpCell.Grid;

/// <summary>Column letters as Excel writes them in headers and addresses.</summary>
internal static class ColumnNames
{
    public static string Of(int column)
    {
        if (column < 1 || column > SheetLayout.MaxColumn)
            throw new ArgumentOutOfRangeException(nameof(column), column, "The column is outside the sheet.");

        Span<char> letters = stackalloc char[3];
        var start = letters.Length;
        while (column > 0)
        {
            column--;
            letters[--start] = (char)('A' + column % 26);
            column /= 26;
        }
        return new string(letters[start..]);
    }
}
