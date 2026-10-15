using System.Globalization;

namespace SharpCell.Grid.Tests.Selection;

/// <summary>Changing how cells look, as Excel's Home tab does, one step to undo each.</summary>
public class FormatTests
{
    private static (Worksheet Sheet, GridEditor Editor) Sheet()
    {
        var workbook = new Workbook { Culture = CultureInfo.GetCultureInfo("en-US") };
        var sheet = workbook.AddSheet("S");
        for (var row = 1; row <= 3; row++)
        {
            for (var column = 1; column <= 3; column++)
                sheet[row, column].Value = row * 10 + column;
        }

        sheet["B2"].Style = CellStyle.Default with { Font = CellFont.Default with { Italic = true } };
        return (sheet, new GridEditor(sheet.Workbook));
    }

    private static CellArea Area(string first, string last)
    {
        var (a, b) = (Cells.Parse(first), Cells.Parse(last));
        return new CellArea(a.Row, a.Column, b.Row, b.Column);
    }

    [Fact]
    public void A_style_change_keeps_the_rest_of_each_style_and_undo_gives_it_back()
    {
        var (sheet, editor) = Sheet();

        editor.SetStyle(sheet, Area("A1", "B2"), style => style with { Font = style.Font with { Bold = true }, Fill = CellColor.FromRgb(0xFFFF00) });

        Assert.True(sheet["B2"].Style.Font.Bold && sheet["B2"].Style.Font.Italic);
        Assert.Equal(CellColor.FromRgb(0xFFFF00), sheet["A1"].Style.Fill);
        Assert.False(sheet["C3"].Style.Font.Bold);

        editor.Undo();
        Assert.Same(sheet.Workbook.DefaultStyle, sheet["A1"].Style);
        Assert.True(sheet["B2"].Style.Font.Italic && !sheet["B2"].Style.Font.Bold);
        editor.Redo();
        Assert.True(sheet["A1"].Style.Font.Bold);
    }

    [Fact]
    public void Number_formats_change_the_text_and_whole_columns_stop_at_the_used_range()
    {
        var (sheet, editor) = Sheet();

        var edit = editor.SetNumberFormat(sheet, new CellArea(1, 2, SheetLayout.MaxRow, 2), "0.00");

        Assert.Equal(new CellArea(1, 2, SheetLayout.MaxRow, 2), edit.Area);
        Assert.Equal("22.00", sheet["B2"].Text);
        Assert.Equal("General", sheet["B4"].NumberFormat);
        editor.Undo();
        Assert.Equal("22", sheet["B2"].Text);
        Assert.Throws<ArgumentException>(() => editor.SetNumberFormat(sheet, Area("A1", "A1"), "0.0.0\""));
        Assert.Equal("General", sheet["A1"].NumberFormat);
    }

    [Fact]
    public void Outside_borders_go_round_the_area_and_no_border_takes_every_line_off()
    {
        var (sheet, editor) = Sheet();
        var thin = new CellBorder(CellBorderStyle.Thin);
        sheet["B1"].Style = CellStyle.Default with { BottomBorder = new CellBorder(CellBorderStyle.Thick) };

        editor.SetBorders(sheet, Area("B2", "C3"), BorderEdges.Outside, thin);

        Assert.Equal((thin, thin, (CellBorder?)null, (CellBorder?)null), (sheet["B2"].Style.TopBorder, sheet["B2"].Style.LeftBorder, sheet["B2"].Style.RightBorder, sheet["B2"].Style.BottomBorder));
        Assert.Equal((thin, thin), (sheet["C3"].Style.RightBorder, sheet["C3"].Style.BottomBorder));
        Assert.Null(sheet["B1"].Style.BottomBorder);                     // the line above is the area's now

        editor.SetBorders(sheet, Area("B2", "C3"), BorderEdges.All, thin);
        Assert.Equal((thin, thin), (sheet["B2"].Style.RightBorder, sheet["C2"].Style.LeftBorder));

        editor.SetBorders(sheet, Area("B2", "C3"), BorderEdges.All, null);
        Assert.Null(sheet["C3"].Style.LeftBorder);
        Assert.True(sheet["B2"].Style.Font.Italic);

        editor.Undo();
        editor.Undo();
        editor.Undo();
        Assert.Equal(CellBorderStyle.Thick, sheet["B1"].Style.BottomBorder!.Style);
    }

    [Theory]
    [InlineData('!', "#,##0.00")]
    [InlineData('%', "0%")]
    [InlineData('$', "$#,##0.00_);($#,##0.00)")]
    [InlineData('~', "General")]
    public void The_ctrl_shift_shortcuts_give_Excel_formats(char symbol, string format)
    {
        Assert.Equal(format, GridEditor.ShortcutFormat(symbol, CultureInfo.GetCultureInfo("en-US")));
        Assert.Equal("#,##0.00 \"€\"", GridEditor.ShortcutFormat('$', CultureInfo.GetCultureInfo("de-DE")));
    }
}
