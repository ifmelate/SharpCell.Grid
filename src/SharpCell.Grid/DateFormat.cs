namespace SharpCell.Grid;

/// <summary>What a number format code shows of a date and time, read from its letters as Excel reads them.</summary>
internal static class DateFormat
{
    /// <summary>Whether the format shows a day, month or year.</summary>
    public static bool ShowsDate(string code) => Letters(code) is var (date, _, _) && date;

    /// <summary>Whether the format shows hours, minutes or seconds and no date.</summary>
    public static bool ShowsTime(string code) => Letters(code) is var (date, time, _) && time && !date;

    // The first section's letters outside quotes, escapes, brackets and the word General; an m among
    // hours or seconds is minutes. [h] and the like are elapsed time.
    private static (bool Date, bool Time, bool Elapsed) Letters(string code)
    {
        bool day = false, year = false, month = false, hour = false, second = false, elapsed = false;
        for (var i = 0; i < code.Length; i++)
        {
            var c = char.ToLowerInvariant(code[i]);
            switch (c)
            {
                case ';':
                    i = code.Length;
                    break;
                case '"':
                    i = code.IndexOf('"', i + 1) is var end and >= 0 ? end : code.Length;
                    break;
                case '\\' or '_' or '*':
                    i++;
                    break;
                case '[':
                    var close = code.IndexOf(']', i + 1);
                    var inside = close < 0 ? "" : code[(i + 1)..close].ToLowerInvariant();
                    if (inside.Length > 0 && inside.Trim('h', 'm', 's').Length == 0)
                        elapsed = hour = true;
                    i = close < 0 ? code.Length : close;
                    break;
                case 'd':
                    day = true;
                    break;
                case 'y':
                    year = true;
                    break;
                case 'g' when string.Compare(code, i, "General", 0, 7, System.StringComparison.OrdinalIgnoreCase) == 0:
                    i += 6;
                    break;
                case 'm':
                    month = true;
                    break;
                case 'h':
                    hour = true;
                    break;
                case 's':
                    second = true;
                    break;
            }
        }

        // m beside h or s is minutes, as in h:mm and mm:ss.
        var minutes = month && (hour || second) && !day && !year;
        var date = day || year || (month && !minutes);
        return (date, hour || second || minutes, elapsed);
    }
}
