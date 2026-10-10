using System;
using System.Globalization;

namespace SharpCell.Grid;

/// <summary>
/// What Excel shows of a number too wide for its column. A number in the General format loses
/// decimals, rounding, or digits of its scientific notation until it fits; any other number, and a
/// General one that cannot shrink enough, shows <c>#</c> across the cell. Text is never changed.
/// </summary>
internal static class TextFitting
{
    // Excel never shows more digits than these in General scientific notation (1.23457E+11).
    private const int MaxMantissaDecimals = 5;

    public static string FitNumber(string text, double value, bool general, CultureInfo culture, double width, Func<string, double> measure)
    {
        if (measure(text) <= width)
            return text;
        if (general && Shorten(text, value, culture, width, measure) is { } shorter)
            return shorter;
        return Hashes(width, measure);
    }

    private static string? Shorten(string text, double value, CultureInfo culture, double width, Func<string, double> measure)
    {
        var separator = culture.NumberFormat.NumberDecimalSeparator;
        if (!text.Contains('E', StringComparison.Ordinal))
        {
            var point = text.IndexOf(separator, StringComparison.Ordinal);
            var decimals = point < 0 ? 0 : text.Length - point - separator.Length;
            for (var keep = decimals - 1; keep >= 0; keep--)
            {
                var rounded = Math.Round(value, keep, MidpointRounding.AwayFromZero);
                var candidate = rounded == 0 ? "0" : rounded.ToString(keep == 0 ? "0" : "0." + new string('#', keep), culture);
                if (measure(candidate) <= width)
                    return candidate;
            }
        }

        for (var keep = MaxMantissaDecimals - 1; keep >= 0; keep--)
        {
            var candidate = value.ToString(keep == 0 ? "0E+00" : "0." + new string('#', keep) + "E+00", culture);
            if (measure(candidate) <= width)
                return candidate;
        }
        return null;
    }

    private static string Hashes(double width, Func<string, double> measure)
    {
        var hash = measure("#");
        var count = hash > 0 ? (int)Math.Floor(width / hash) : 1;
        return new string('#', Math.Max(1, count));
    }
}
