using System.Globalization;

namespace SharpCell.Grid.Tests.Selection;

public class FillTests
{
    private static (Worksheet Sheet, GridEditor Editor) Sheet()
    {
        var workbook = new Workbook { Culture = CultureInfo.GetCultureInfo("en-US") };
        return (workbook.AddSheet("S"), new GridEditor(workbook));
    }

    private static CellArea Area(string first, string? last = null)
    {
        var (a, b) = (Cells.Parse(first), Cells.Parse(last ?? first));
        return new CellArea(a.Row, a.Column, b.Row, b.Column);
    }

    private static string[] Texts(Worksheet sheet, params string[] addresses) => addresses.Select(a => sheet[a].Text).ToArray();

    [Fact]
    public void A_single_number_is_copied_and_two_numbers_continue_their_step()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "5";
        sheet["B1"].Input = "1";
        sheet["B2"].Input = "3";

        editor.Fill(sheet, Area("A1"), Area("A1", "A3"));
        editor.Fill(sheet, Area("B1", "B2"), Area("B1", "B5"));

        Assert.Equal(["5", "5"], Texts(sheet, "A2", "A3"));
        Assert.Equal(["5", "7", "9"], Texts(sheet, "B3", "B4", "B5"));
    }

    [Fact]
    public void Uneven_numbers_continue_their_trend_as_in_Excel()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "1";
        sheet["A2"].Input = "2";
        sheet["A3"].Input = "4";

        editor.Fill(sheet, Area("A1", "A3"), Area("A1", "A5"));

        // Excel's linear trend through (0, 1), (1, 2), (2, 4): 5/6 + 1.5x.
        Assert.Equal(5.3333333333333339, sheet["A4"].Value.AsNumber(), 12);
        Assert.Equal(6.8333333333333339, sheet["A5"].Value.AsNumber(), 12);
    }

    [Fact]
    public void Filling_up_or_left_continues_the_series_backwards()
    {
        var (sheet, editor) = Sheet();
        sheet["A3"].Input = "10";
        sheet["A4"].Input = "20";
        sheet["E1"].Input = "Item 5";

        editor.Fill(sheet, Area("A3", "A4"), Area("A1", "A4"));
        editor.Fill(sheet, Area("E1"), Area("C1", "E1"));

        Assert.Equal(["-10", "0"], Texts(sheet, "A1", "A2"));
        Assert.Equal(["Item 3", "Item 4"], Texts(sheet, "C1", "D1"));
    }

    [Fact]
    public void Formulas_are_copied_with_their_references_following()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "1";
        sheet["A2"].Input = "2";
        sheet["A3"].Input = "3";
        sheet["B1"].Input = "=A1*$A$1*10";
        sheet.Workbook.Recalculate();

        editor.Fill(sheet, Area("B1"), Area("B1", "B3"));

        Assert.Equal("=A3*$A$1*10", sheet["B3"].Formula);
        Assert.Equal("30", sheet["B3"].Text);
    }

    [Fact]
    public void A_single_date_steps_a_day_and_a_time_an_hour()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "1/31/2026";
        sheet["B1"].Input = "9:00";

        editor.Fill(sheet, Area("A1", "B1"), Area("A1", "B3"));

        Assert.Equal(["2/1/2026", "2/2/2026"], Texts(sheet, "A2", "A3"));
        Assert.Equal(["10:00", "11:00"], Texts(sheet, "B2", "B3"));
    }

    [Fact]
    public void Dates_a_month_apart_keep_their_day_and_fall_back_to_the_end_of_short_months()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "11/30/2025";
        sheet["A2"].Input = "12/30/2025";

        editor.Fill(sheet, Area("A1", "A2"), Area("A1", "A4"));

        Assert.Equal(["1/30/2026", "2/28/2026"], Texts(sheet, "A3", "A4"));
    }

    [Fact]
    public void Text_ending_in_a_number_counts_on_with_its_zeros()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "Item 1";
        sheet["B1"].Input = "A007";
        sheet["C1"].Input = "Week 2";
        sheet["C2"].Input = "Week 4";

        editor.Fill(sheet, Area("A1", "B1"), Area("A1", "B3"));
        editor.Fill(sheet, Area("C1", "C2"), Area("C1", "C4"));

        Assert.Equal(["Item 2", "Item 3"], Texts(sheet, "A2", "A3"));
        Assert.Equal(["A008", "A009"], Texts(sheet, "B2", "B3"));
        Assert.Equal(["Week 6", "Week 8"], Texts(sheet, "C3", "C4"));
    }

    [Fact]
    public void Day_and_month_names_continue_in_their_form_and_case()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "Friday";
        sheet["B1"].Input = "NOV";
        sheet["C1"].Input = "mon";
        sheet["C2"].Input = "wed";

        editor.Fill(sheet, Area("A1", "B1"), Area("A1", "B4"));
        editor.Fill(sheet, Area("C1", "C2"), Area("C1", "C4"));

        Assert.Equal(["Saturday", "Sunday", "Monday"], Texts(sheet, "A2", "A3", "A4"));
        Assert.Equal(["DEC", "JAN", "FEB"], Texts(sheet, "B2", "B3", "B4"));
        Assert.Equal(["fri", "sun"], Texts(sheet, "C3", "C4"));
    }

    [Fact]
    public void Other_cells_repeat_with_their_formats_and_styles()
    {
        var (sheet, editor) = Sheet();
        var bold = CellStyle.Default with { Font = CellFont.Default with { Bold = true } };
        sheet["A1"].Input = "North";
        sheet["A1"].Style = bold;
        sheet["A2"].Input = "TRUE";

        editor.Fill(sheet, Area("A1", "A2"), Area("A1", "A5"));

        Assert.Equal(["North", "TRUE", "North"], Texts(sheet, "A3", "A4", "A5"));
        Assert.Equal(bold, sheet["A3"].Style);
        Assert.Equal(bold, sheet["A5"].Style);
    }

    [Fact]
    public void Dragging_the_handle_back_clears_the_cells_left_out()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "1";
        sheet["A2"].Input = "2";
        sheet["A3"].Input = "3";

        editor.Fill(sheet, Area("A1", "A3"), Area("A1"));

        Assert.Equal(["1", "", ""], Texts(sheet, "A1", "A2", "A3"));
    }

    [Fact]
    public void A_fill_is_one_step_to_undo()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].Input = "Item 1";
        sheet["A3"].Input = "keep";
        sheet["A3"].NumberFormat = "@";

        editor.Fill(sheet, Area("A1"), Area("A1", "A4"));
        editor.Undo();

        Assert.Equal(["Item 1", "", "keep", ""], Texts(sheet, "A1", "A2", "A3", "A4"));
        Assert.Equal("@", sheet["A3"].NumberFormat);
    }

    [Fact]
    public void A_fill_not_reaching_past_the_cells_on_one_side_is_refused()
    {
        var (sheet, editor) = Sheet();

        Assert.Throws<ArgumentException>(() => editor.Fill(sheet, Area("A1", "B2"), Area("A1", "C3")));
        Assert.Throws<ArgumentException>(() => editor.Fill(sheet, Area("A1", "B2"), Area("A2", "B3")));
        Assert.False(editor.CanUndo);
    }

    [Theory]
    [InlineData("B5", "B2", "C5")]
    [InlineData("D3", "B2", "D3")]
    [InlineData("A1", "B1", "C3")]
    [InlineData("F4", "B2", "F3")]
    [InlineData("B2", "B2", "C2")]
    [InlineData("C2", "B2", "C2")]
    [InlineData("B3", "B2", "B3")]
    public void Dragging_the_handle_picks_the_side_the_pointer_is_furthest_beyond(string pointer, string first, string last)
    {
        Assert.Equal(Area(first, last), GridEditor.FillArea(Area("B2", "C3"), Cells.Parse(pointer)));
    }
}
