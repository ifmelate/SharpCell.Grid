using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace SharpCell.Grid;

/// <summary>
/// What Excel's fill handle puts in each cell it fills: which cell is copied there, and the value
/// that takes the place of the copy's when the cells continue a series. Each column filled down or
/// up, or each row filled right or left, is a line of its own.
/// </summary>
internal static class FillSeries
{
    /// <summary>A filled cell: copied from <see cref="From"/>, with <see cref="Value"/> instead of the copy's when it is a series.</summary>
    internal readonly record struct Filled(CellPosition From, CellPosition To, CellValue? Value);

    public static List<Filled> Plan(Worksheet sheet, CellArea from, CellArea to)
    {
        var down = to.LastRow > from.LastRow;
        var up = to.FirstRow < from.FirstRow;
        var right = to.LastColumn > from.LastColumn;
        var left = to.FirstColumn < from.FirstColumn;
        var sameColumns = to.FirstColumn == from.FirstColumn && to.LastColumn == from.LastColumn;
        var sameRows = to.FirstRow == from.FirstRow && to.LastRow == from.LastRow;
        if (!(((down ^ up) && sameColumns && (down ? to.FirstRow == from.FirstRow : to.LastRow == from.LastRow))
            || ((right ^ left) && sameRows && (right ? to.FirstColumn == from.FirstColumn : to.LastColumn == from.LastColumn))))
            throw new ArgumentException("The cells filled must hold the cells filled from and reach past them on one side.", nameof(to));

        var plan = new List<Filled>();
        var vertical = down || up;
        var forward = down || right;
        var (lines, length) = vertical
            ? (Enumerable.Range(from.FirstColumn, from.LastColumn - from.FirstColumn + 1), from.LastRow - from.FirstRow + 1)
            : (Enumerable.Range(from.FirstRow, from.LastRow - from.FirstRow + 1), from.LastColumn - from.FirstColumn + 1);
        var count = vertical ? to.LastRow - to.FirstRow + 1 - length : to.LastColumn - to.FirstColumn + 1 - length;
        foreach (var line in lines)
        {
            // Cells in the order the fill goes: away from the cells filled from.
            CellPosition At(int step)
            {
                var offset = forward ? step : -step;
                return vertical
                    ? new CellPosition(forward ? from.FirstRow + offset : from.LastRow + offset, line)
                    : new CellPosition(line, forward ? from.FirstColumn + offset : from.LastColumn + offset);
            }

            var source = Enumerable.Range(0, length).Select(i => sheet[At(i).Row, At(i).Column]).ToList();
            var values = Continue(source, count, forward, sheet.Workbook);
            for (var i = 0; i < count; i++)
                plan.Add(new Filled(At(i % length), At(length + i), values[i]));
        }

        return plan;
    }

    // The values of the cells after the line's cells, null where the cell is copied as it is.
    private static CellValue?[] Continue(List<Cell> cells, int count, bool forward, Workbook workbook)
    {
        var values = new CellValue?[count];
        var n = cells.Count;
        var sign = forward ? 1 : -1;
        var items = cells.Select(c => Item.Of(c, workbook)).ToList();

        // A line of numbers continues its trend; dates a month or a year apart keep their day.
        if (items.All(i => i.Kind is ItemKind.Number or ItemKind.Date or ItemKind.Time) && (n > 1 || items[0].Kind != ItemKind.Number))
        {
            if (n == 1)
            {
                var step = items[0].Kind == ItemKind.Time ? 1.0 / 24 : 1;
                for (var i = 0; i < count; i++)
                    values[i] = CellValue.Number(items[0].Number + sign * step * (i + 1));
                return values;
            }

            var numbers = items.Select(i => i.Number).ToArray();
            if (items.All(i => i.Kind == ItemKind.Date) && Months(numbers, workbook.DateSystem) is { } months)
            {
                for (var i = 0; i < count; i++)
                    values[i] = CellValue.Number(AddMonths(numbers[0], months * (n + i), workbook.DateSystem));
                return values;
            }

            var (a, b) = Trend(numbers);
            for (var i = 0; i < count; i++)
                values[i] = CellValue.Number(a + b * (n + i));
            return values;
        }

        // Otherwise the cells repeat. Those of one kind, as text with the same words before a number
        // or names of days, continue together; each alone steps by one.
        var groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        for (var i = 0; i < n; i++)
        {
            if (items[i].Group is { } key)
            {
                if (!groups.TryGetValue(key, out var members))
                    groups[key] = members = [];
                members.Add(i);
            }
        }

        for (var i = 0; i < count; i++)
        {
            var position = i % n;
            var cycle = i / n + 1;
            var item = items[position];
            if (item.Group is not { } key)
                continue;
            var members = groups[key];
            var m = members.Count;
            double step;
            if (m == 1)
                step = item.Kind == ItemKind.Number ? 0 : item.Kind == ItemKind.Time ? sign / 24.0 : sign;
            else if (item.Kind == ItemKind.Name)
                step = ((items[members[1]].Number - items[members[0]].Number) % item.Names!.Count + item.Names.Count) % item.Names.Count;
            else
                step = (items[members[^1]].Number - items[members[0]].Number) / (m - 1);
            var next = item.Number + step * m * cycle;
            values[i] = item.Kind switch
            {
                ItemKind.Text => CellValue.Text(item.Prefix + Math.Abs(Math.Round(next)).ToString(CultureInfo.InvariantCulture).PadLeft(item.Digits, '0')),
                ItemKind.Name => CellValue.Text(item.Names!.Name(next, item.Case)),
                _ => step == 0 ? (CellValue?)null : CellValue.Number(next),
            };
        }

        return values;
    }

    // The line through the points (0, v0), (1, v1)...: a and b of a + b * x. Evenly spaced numbers
    // keep their exact step.
    private static (double A, double B) Trend(double[] values)
    {
        var n = values.Length;
        var step = (values[^1] - values[0]) / (n - 1);
        var even = true;
        for (var i = 1; i < n; i++)
            even &= Math.Abs(values[i] - values[i - 1] - step) <= 1e-12 * Math.Max(1, Math.Abs(step));
        if (even)
            return (values[0], step);

        var meanX = (n - 1) / 2.0;
        var meanY = values.Average();
        double sxy = 0, sxx = 0;
        for (var i = 0; i < n; i++)
        {
            sxy += (i - meanX) * (values[i] - meanY);
            sxx += (i - meanX) * (i - meanX);
        }

        var b = sxy / sxx;
        return (meanY - b * meanX, b);
    }

    // Whole dates on the same day of the month, the same number of months apart: that number.
    private static int? Months(double[] serials, DateSystem system)
    {
        var dates = new List<DateTime>();
        foreach (var serial in serials)
        {
            if (serial != Math.Floor(serial) || ToDate(serial, system) is not { } date)
                return null;
            dates.Add(date);
        }

        if (dates.Any(d => d.Day != dates[0].Day))
            return null;
        int MonthOf(DateTime d) => d.Year * 12 + d.Month;
        var step = MonthOf(dates[1]) - MonthOf(dates[0]);
        if (step == 0)
            return null;
        for (var i = 2; i < dates.Count; i++)
        {
            if (MonthOf(dates[i]) - MonthOf(dates[i - 1]) != step)
                return null;
        }

        return step;
    }

    // As EDATE: a day past the end of the month falls back to its last day.
    private static double AddMonths(double serial, int months, DateSystem system) =>
        ToDate(serial, system) is { } date && FromDate(date.AddMonths(months), system) is { } next ? next : serial;

    // Serials from 1 March 1900 on, past Excel's 29 February 1900, in the 1900 system.
    private static DateTime? ToDate(double serial, DateSystem system)
    {
        var origin = system == DateSystem.Date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
        if (system != DateSystem.Date1904 && serial < 61)
            return null;
        var date = origin.AddDays(serial);
        return date.Year is >= 1900 and <= 9999 ? date : null;
    }

    private static double? FromDate(DateTime date, DateSystem system)
    {
        var origin = system == DateSystem.Date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
        return (date - origin).TotalDays;
    }

    private enum ItemKind
    {
        Other,
        Number,
        Date,
        Time,
        Text,
        Name,
    }

    // A cell as the fill sees it: a number, a date or time, text ending in a number, the name of a
    // day or month, or something that is only copied. Cells of one group continue together.
    private sealed record Item(ItemKind Kind, double Number, string? Group, string Prefix = "", int Digits = 0, NameList? Names = null, LetterCase Case = LetterCase.AsListed)
    {
        public static Item Of(Cell cell, Workbook workbook)
        {
            if (cell.Formula is not null)
                return new Item(ItemKind.Other, 0, null);
            var value = cell.Value;
            if (value.Kind == CellValueKind.Number)
            {
                var format = cell.NumberFormat;
                var kind = DateFormat.ShowsDate(format) ? ItemKind.Date : DateFormat.ShowsTime(format) ? ItemKind.Time : ItemKind.Number;
                return new Item(kind, value.AsNumber(), kind switch { ItemKind.Date => "date", ItemKind.Time => "time", _ => "number" });
            }

            if (value.Kind != CellValueKind.Text)
                return new Item(ItemKind.Other, 0, null);
            var text = value.AsText();
            foreach (var list in NameList.For(workbook.Culture))
            {
                if (list.IndexOf(text) is { } index)
                    return new Item(ItemKind.Name, index, "names:" + list.Key, Names: list, Case: LetterCaseOf(text));
            }

            // Text ending in a number: the number continues, zeros before it kept.
            var digits = 0;
            while (digits < text.Length && char.IsAsciiDigit(text[^(digits + 1)]))
                digits++;
            if (digits == 0 || digits > 15)
                return new Item(ItemKind.Other, 0, null);
            var prefix = text[..^digits];
            var number = double.Parse(text.AsSpan(text.Length - digits), NumberStyles.None, CultureInfo.InvariantCulture);
            return new Item(ItemKind.Text, number, "text:" + prefix, prefix, digits);
        }
    }

    private enum LetterCase
    {
        AsListed,
        Upper,
        Lower,
    }

    private static LetterCase LetterCaseOf(string text) =>
        text.Any(char.IsLetter) && text == text.ToUpperInvariant() ? LetterCase.Upper
        : text == text.ToLowerInvariant() ? LetterCase.Lower
        : LetterCase.AsListed;

    // Names of days or months, full or short, in a culture; English ones too, as Excel knows them.
    private sealed class NameList(string key, string[] names, CultureInfo culture)
    {
        public string Key => key;

        public int Count => names.Length;

        public static IEnumerable<NameList> For(CultureInfo culture)
        {
            foreach (var c in culture.Name.Length == 0 || culture.TwoLetterISOLanguageName == "en"
                ? [CultureInfo.InvariantCulture] : new[] { culture, CultureInfo.InvariantCulture })
            {
                var format = c.DateTimeFormat;
                yield return new NameList(c.Name + ":days", format.DayNames, c);
                yield return new NameList(c.Name + ":shortdays", format.AbbreviatedDayNames, c);
                yield return new NameList(c.Name + ":months", format.MonthNames[..12], c);
                yield return new NameList(c.Name + ":shortmonths", format.AbbreviatedMonthNames[..12], c);
            }
        }

        public int? IndexOf(string text)
        {
            var trimmed = text.Trim();
            for (var i = 0; i < names.Length; i++)
            {
                if (names[i].Length > 0 && string.Equals(names[i], trimmed, StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            return null;
        }

        public string Name(double index, LetterCase letterCase)
        {
            var at = (int)(((long)Math.Round(index) % names.Length + names.Length) % names.Length);
            var name = names[at];
            return letterCase switch
            {
                LetterCase.Upper => culture.TextInfo.ToUpper(name),
                LetterCase.Lower => culture.TextInfo.ToLower(name),
                _ => name,
            };
        }
    }
}
