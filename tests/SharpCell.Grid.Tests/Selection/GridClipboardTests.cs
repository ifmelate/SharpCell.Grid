using System.Text;

namespace SharpCell.Grid.Tests.Selection;

public class GridClipboardTests
{
    private static CellArea Area(string first, string last)
    {
        var (a, b) = (Cells.Parse(first), Cells.Parse(last));
        return new CellArea(a.Row, a.Column, b.Row, b.Column);
    }

    private static Worksheet Sheet()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = "Item";
        sheet["B1"].Value = "Price";
        sheet["A2"].Value = "Rent";
        sheet["B2"].Value = 1250;
        sheet["B2"].NumberFormat = "#,##0.00";
        sheet["A3"].Value = "Two\nlines";
        sheet["B3"].Formula = "=B2*2";
        sheet["B3"].NumberFormat = "#,##0.00";
        workbook.Recalculate();
        return sheet;
    }

    [Fact]
    public void Text_has_tabs_between_cells_and_a_line_break_after_every_row()
    {
        Assert.Equal("Item\tPrice\r\nRent\t1,250.00\r\n", GridClipboard.Text(Sheet(), Area("A1", "B2")));
    }

    [Fact]
    public void Formulas_are_copied_as_their_shown_values()
    {
        Assert.Equal("2,500.00\r\n", GridClipboard.Text(Sheet(), Area("B3", "B3")));
    }

    [Fact]
    public void Text_with_a_line_break_is_quoted()
    {
        var sheet = Sheet();
        sheet["C3"].Value = "say \"hi\"\tnow";

        Assert.Equal("\"Two\nlines\"\t2,500.00\t\"say \"\"hi\"\"\tnow\"\r\n", GridClipboard.Text(sheet, Area("A3", "C3")));
    }

    [Fact]
    public void Empty_cells_are_copied_as_empty()
    {
        Assert.Equal("\t\r\n\t\r\n", GridClipboard.Text(Sheet(), Area("D1", "E2")));
    }

    [Fact]
    public void Whole_columns_and_rows_end_at_the_used_range()
    {
        var sheet = Sheet();

        Assert.Equal("Item\r\nRent\r\n\"Two\nlines\"\r\n", GridClipboard.Text(sheet, new CellArea(1, 1, SheetLayout.MaxRow, 1)));
        Assert.Equal("Item\tPrice\r\n", GridClipboard.Text(sheet, new CellArea(1, 1, 1, SheetLayout.MaxColumn)));
        Assert.Equal("\r\n", GridClipboard.Text(new Workbook().AddSheet("Empty"), new CellArea(1, 1, SheetLayout.MaxRow, SheetLayout.MaxColumn)));
    }

    [Fact]
    public void Html_is_a_table_with_the_cells_as_shown()
    {
        var html = GridClipboard.Html(Sheet(), Area("A1", "B3"));

        Assert.Contains("<!--StartFragment--><table", html);
        Assert.Contains("</table><!--EndFragment-->", html);
        Assert.Equal(3, html.Split("<tr>").Length - 1);
        Assert.Contains(">1,250.00</td>", html);
        Assert.Contains(">Two<br style=\"mso-data-placement:same-cell\">lines</td>", html);
    }

    [Fact]
    public void Html_keeps_styles_and_keeps_text_as_text()
    {
        var sheet = Sheet();
        sheet["A1"].Style = CellStyle.Default with
        {
            Font = CellFont.Default with { Bold = true, Color = CellColor.FromRgb(0xFF0000) },
            Fill = CellColor.FromRgb(0xD9E1F2),
            HorizontalAlignment = CellHorizontalAlignment.Center,
            BottomBorder = new CellBorder(CellBorderStyle.Thin, null),
        };

        var html = GridClipboard.Html(sheet, Area("A1", "B1"));

        var item = html[html.IndexOf("<td", StringComparison.Ordinal)..html.IndexOf(">Item<", StringComparison.Ordinal)];
        Assert.Contains("font-weight:700", item);
        Assert.Contains("color:#FF0000", item);
        Assert.Contains("background:#D9E1F2", item);
        Assert.Contains("text-align:center", item);
        Assert.Contains("border-bottom:.5pt solid windowtext", item);
        Assert.Contains("mso-number-format:\"\\@\"", item);
        Assert.DoesNotContain("mso-number-format", GridClipboard.Html(sheet, Area("B2", "B2")));
    }

    [Fact]
    public void Html_escapes_markup_in_cells()
    {
        var sheet = Sheet();
        sheet["A1"].Value = "<b>&'";

        Assert.Contains(">&lt;b&gt;&amp;&#39;</td>", GridClipboard.Html(sheet, Area("A1", "A1")));
    }

    [Fact]
    public void A_merged_area_is_one_cell_spanning_its_columns_and_rows()
    {
        var sheet = Sheet();
        sheet["D1"].Value = "Title";
        sheet.Merge("D1:E2");

        var html = GridClipboard.Html(sheet, Area("C1", "E2"));

        Assert.Contains("colspan=\"2\" rowspan=\"2\"", html);
        Assert.Equal(1, html.Split(">Title<").Length - 1);
        Assert.Equal(3, html.Split("<td").Length - 1); // C1, D1:E2, C2
    }

    [Fact]
    public void Windows_html_gives_byte_offsets_of_the_document_and_its_fragment()
    {
        var html = GridClipboard.Html(Sheet(), Area("A1", "B2")).Replace("Item", "Prix €", StringComparison.Ordinal);

        var wrapped = Encoding.UTF8.GetBytes(GridClipboard.WindowsHtml(html));

        var header = Encoding.ASCII.GetString(wrapped, 0, 200);
        int Offset(string name) => int.Parse(header.Split(name + ":")[1][..10]);
        Assert.StartsWith("Version:0.9\r\n", header);
        Assert.Equal("<html", Encoding.UTF8.GetString(wrapped, Offset("StartHTML"), 5));
        Assert.Equal(wrapped.Length, Offset("EndHTML"));
        var fragment = Encoding.UTF8.GetString(wrapped, Offset("StartFragment"), Offset("EndFragment") - Offset("StartFragment"));
        Assert.StartsWith("<table", fragment);
        Assert.EndsWith("</table>", fragment);
        Assert.Contains("Prix €", fragment);
    }

    [Fact]
    public void Windows_html_needs_the_fragment_comments()
    {
        Assert.Throws<ArgumentException>(() => GridClipboard.WindowsHtml("<html><table></table></html>"));
    }
}
