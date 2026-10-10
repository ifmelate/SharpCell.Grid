using System;
using System.Globalization;
using System.Net;
using System.Text;

namespace SharpCell.Grid;

/// <summary>
/// What copying cells puts on the clipboard, in the forms Excel reads when it pastes: text with
/// tabs between cells, and an HTML table that keeps fonts, fills, borders, alignment and merged
/// cells. Cells are copied as shown, <see cref="Cell.Text"/> in their number format. Hidden rows and
/// columns are copied, as Excel copies them. Whole columns or rows end at the sheet's used range.
/// </summary>
public static class GridClipboard
{
    /// <summary>
    /// The cells as text, as Excel puts them on the clipboard: a tab between cells, a line break
    /// after each row, the last one too. Text with a line break or a tab is put in double quotes,
    /// with any quote inside doubled.
    /// </summary>
    public static string Text(Worksheet sheet, CellArea area)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        area = Copied(sheet, area);
        var text = new StringBuilder();
        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
            {
                if (column > area.FirstColumn)
                    text.Append('\t');
                var shown = sheet[row, column].Text;
                if (shown.AsSpan().IndexOfAny('\n', '\r', '\t') >= 0)
                    text.Append('"').Append(shown.Replace("\"", "\"\"", StringComparison.Ordinal)).Append('"');
                else
                    text.Append(shown);
            }

            text.Append("\r\n");
        }

        return text.ToString();
    }

    /// <summary>
    /// The cells as an HTML document with one table, whose cells carry their style inline and merged
    /// areas as <c>colspan</c> and <c>rowspan</c>; the table lies between the
    /// <c>&lt;!--StartFragment--&gt;</c> and <c>&lt;!--EndFragment--&gt;</c> comments. On macOS and
    /// Linux it goes on the clipboard as it is; Windows needs it wrapped by <see cref="WindowsHtml"/>.
    /// </summary>
    public static string Html(Worksheet sheet, CellArea area)
    {
        ArgumentNullException.ThrowIfNull(sheet);
        area = Copied(sheet, area);
        var html = new StringBuilder();
        html.Append("<html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\">\r\n");
        html.Append("<head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head>\r\n<body>\r\n");
        html.Append("<!--StartFragment-->");
        html.Append("<table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" style=\"border-collapse:collapse\">\r\n");
        var merges = sheet.MergedAreas;
        for (var row = area.FirstRow; row <= area.LastRow; row++)
        {
            html.Append("<tr>");
            for (var column = area.FirstColumn; column <= area.LastColumn; column++)
            {
                // A merged area is one cell, cut to the copied cells, showing its top-left cell.
                var (span, source) = (new CellArea(row, column, row, column), sheet[row, column]);
                if (GridNavigation.MergeAt(merges, new CellPosition(row, column)) is { } merge)
                {
                    var (top, left) = (Math.Max(merge.FirstRow, area.FirstRow), Math.Max(merge.FirstColumn, area.FirstColumn));
                    if (top != row || left != column)
                        continue;
                    span = new CellArea(top, left, Math.Min(merge.LastRow, area.LastRow), Math.Min(merge.LastColumn, area.LastColumn));
                    source = sheet[merge.FirstRow, merge.FirstColumn];
                }

                AppendCell(html, sheet, source, span);
            }

            html.Append("</tr>\r\n");
        }

        html.Append("</table>");
        html.Append("<!--EndFragment-->\r\n</body>\r\n</html>");
        return html.ToString();
    }

    /// <summary>
    /// HTML in the clipboard format Windows applications read, <c>HTML Format</c>: a header giving
    /// where the document and its fragment start and end, as byte offsets into its UTF-8 encoding,
    /// then the document. Put it on the clipboard as UTF-8 bytes.
    /// </summary>
    /// <exception cref="ArgumentException">The document has no <c>&lt;!--StartFragment--&gt;</c> and <c>&lt;!--EndFragment--&gt;</c> comments.</exception>
    public static string WindowsHtml(string html)
    {
        ArgumentNullException.ThrowIfNull(html);
        const string Start = "<!--StartFragment-->";
        const string End = "<!--EndFragment-->";
        var startIndex = html.IndexOf(Start, StringComparison.Ordinal);
        var endIndex = html.IndexOf(End, StringComparison.Ordinal);
        if (startIndex < 0 || endIndex < startIndex)
            throw new ArgumentException("The HTML has no fragment between <!--StartFragment--> and <!--EndFragment-->.", nameof(html));

        // Ten digits for each offset keep the header the same length whatever the offsets are.
        const string Header = "Version:0.9\r\nStartHTML:{0:D10}\r\nEndHTML:{1:D10}\r\nStartFragment:{2:D10}\r\nEndFragment:{3:D10}\r\n";
        var headerLength = string.Format(CultureInfo.InvariantCulture, Header, 0, 0, 0, 0).Length;
        var utf8 = Encoding.UTF8;
        var startFragment = headerLength + utf8.GetByteCount(html.AsSpan(0, startIndex + Start.Length));
        var endFragment = headerLength + utf8.GetByteCount(html.AsSpan(0, endIndex));
        var endHtml = headerLength + utf8.GetByteCount(html);
        return string.Format(CultureInfo.InvariantCulture, Header, headerLength, endHtml, startFragment, endFragment) + html;
    }

    // The copied cells: those asked for, whole columns and rows cut at the used range.
    private static CellArea Copied(Worksheet sheet, CellArea area)
    {
        if (area.FirstRow < 1 || area.FirstColumn < 1 || area.LastRow > SheetLayout.MaxRow || area.LastColumn > SheetLayout.MaxColumn || area.LastRow < area.FirstRow || area.LastColumn < area.FirstColumn)
            throw new ArgumentOutOfRangeException(nameof(area), area, "The area is not a rectangle of cells inside the sheet.");
        var used = sheet.UsedRange;
        var (lastRow, lastColumn) = (area.LastRow, area.LastColumn);
        if (area.FirstRow == 1 && area.LastRow == SheetLayout.MaxRow)
            lastRow = Math.Max(area.FirstRow, used?.LastRow ?? 1);
        if (area.FirstColumn == 1 && area.LastColumn == SheetLayout.MaxColumn)
            lastColumn = Math.Max(area.FirstColumn, used?.LastColumn ?? 1);
        return area with { LastRow = lastRow, LastColumn = lastColumn };
    }

    private static void AppendCell(StringBuilder html, Worksheet sheet, Cell cell, CellArea span)
    {
        var style = cell.Style;
        var font = style.Font;
        var css = new StringBuilder();
        css.Append("font-family:").Append(CssString(font.Name)).Append(';');
        css.Append("font-size:").Append(font.Size.ToString("0.##", CultureInfo.InvariantCulture)).Append("pt;");
        css.Append("font-weight:").Append(font.Bold ? "700" : "400").Append(';');
        css.Append("font-style:").Append(font.Italic ? "italic" : "normal").Append(';');
        var decoration = (font.Underline != CellUnderline.None, font.Strikethrough) switch
        {
            (true, true) => "underline line-through",
            (true, false) => "underline",
            (false, true) => "line-through",
            _ => "none",
        };
        css.Append("text-decoration:").Append(decoration).Append(';');
        css.Append("color:").Append(font.Color?.ToString() ?? "#000000").Append(';');
        if (style.Fill is { } fill)
            css.Append("background:").Append(fill).Append(';');
        var align = style.HorizontalAlignment switch
        {
            CellHorizontalAlignment.Left => "left",
            CellHorizontalAlignment.Center or CellHorizontalAlignment.CenterContinuous => "center",
            CellHorizontalAlignment.Right => "right",
            CellHorizontalAlignment.Justify or CellHorizontalAlignment.Distributed => "justify",
            _ => null,
        };
        if (align is not null)
            css.Append("text-align:").Append(align).Append(';');
        css.Append("vertical-align:").Append(style.VerticalAlignment switch
        {
            CellVerticalAlignment.Top => "top",
            CellVerticalAlignment.Center or CellVerticalAlignment.Justify or CellVerticalAlignment.Distributed => "middle",
            _ => "bottom",
        }).Append(';');
        css.Append("white-space:").Append(style.WrapText ? "normal" : "nowrap").Append(';');
        if (style.Indent > 0)
            css.Append("padding-left:").Append((style.Indent * 9).ToString(CultureInfo.InvariantCulture)).Append("pt;");

        // Each merged cell's edges come from the cells on its outline, as Excel draws them.
        Border(css, "border-left", sheet[span.FirstRow, span.FirstColumn].Style.LeftBorder);
        Border(css, "border-right", sheet[span.FirstRow, span.LastColumn].Style.RightBorder);
        Border(css, "border-top", sheet[span.FirstRow, span.FirstColumn].Style.TopBorder);
        Border(css, "border-bottom", sheet[span.LastRow, span.FirstColumn].Style.BottomBorder);

        // Text stays text in Excel, even when it reads as a number or a date.
        if (cell.Value.Kind == CellValueKind.Text)
            css.Append("mso-number-format:\"\\@\";");

        html.Append("<td");
        if (span.LastColumn > span.FirstColumn)
            html.Append(" colspan=\"").Append((span.LastColumn - span.FirstColumn + 1).ToString(CultureInfo.InvariantCulture)).Append('"');
        if (span.LastRow > span.FirstRow)
            html.Append(" rowspan=\"").Append((span.LastRow - span.FirstRow + 1).ToString(CultureInfo.InvariantCulture)).Append('"');
        html.Append(" style='").Append(css.ToString().Replace("'", "&#39;", StringComparison.Ordinal)).Append("'>");

        // A line break inside a cell stays in the cell when Excel pastes it.
        var lines = cell.Text.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            if (i > 0)
                html.Append("<br style=\"mso-data-placement:same-cell\">");
            html.Append(WebUtility.HtmlEncode(lines[i]));
        }

        html.Append("</td>");
    }

    private static void Border(StringBuilder css, string side, CellBorder? border)
    {
        if (border is null)
            return;
        var (width, line) = border.Style switch
        {
            CellBorderStyle.Medium => ("1pt", "solid"),
            CellBorderStyle.Thick => ("1.5pt", "solid"),
            CellBorderStyle.Double => ("2pt", "double"),
            CellBorderStyle.Hair or CellBorderStyle.Dotted => (".5pt", "dotted"),
            CellBorderStyle.Dashed or CellBorderStyle.DashDot or CellBorderStyle.DashDotDot => (".5pt", "dashed"),
            CellBorderStyle.MediumDashed or CellBorderStyle.MediumDashDot or CellBorderStyle.MediumDashDotDot or CellBorderStyle.SlantDashDot => ("1pt", "dashed"),
            _ => (".5pt", "solid"),
        };
        css.Append(side).Append(':').Append(width).Append(' ').Append(line).Append(' ').Append(border.Color?.ToString() ?? "windowtext").Append(';');
    }

    private static string CssString(string value) => "\"" + value.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal) + "\"";
}
