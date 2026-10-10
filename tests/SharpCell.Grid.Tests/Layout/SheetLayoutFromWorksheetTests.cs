namespace SharpCell.Grid.Tests.Layout;

/// <summary>
/// Sizes in pixels from a worksheet's widths and heights, as Excel converts them at 100% zoom with
/// Calibri 11, whose widest digit is 7 pixels.
/// </summary>
public class SheetLayoutFromWorksheetTests
{
    private static Worksheet Sheet() => new Workbook().AddSheet("S");

    [Fact]
    public void A_sheet_without_sizes_has_Excels_defaults()
    {
        var layout = SheetLayout.FromWorksheet(Sheet(), maxDigitWidth: 7);

        Assert.Equal(64, layout.ColumnWidth(1));
        Assert.Equal(20, layout.RowHeight(1));
    }

    [Theory]
    [InlineData(9.140625, 64)]       // Excel's 8.43 characters
    [InlineData(21.42578125, 150)]   // 20.71 characters
    [InlineData(0.7109375, 5)]       // 0 characters: the padding alone
    [InlineData(0, 0)]
    public void Column_widths_are_converted_as_Excel_does(double width, double pixels)
    {
        var sheet = Sheet();
        sheet.SetColumnWidth(2, width);

        Assert.Equal(pixels, SheetLayout.FromWorksheet(sheet, 7).ColumnWidth(2));
    }

    [Fact]
    public void The_default_width_comes_from_the_sheet()
    {
        var sheet = Sheet();
        sheet.DefaultColumnWidth = 21.42578125;

        Assert.Equal(150, SheetLayout.FromWorksheet(sheet, 7).ColumnWidth(9));
    }

    [Theory]
    [InlineData(15, 20)]
    [InlineData(30, 40)]
    [InlineData(14.4, 19)]
    public void Row_heights_in_points_become_whole_pixels(double points, double pixels)
    {
        var sheet = Sheet();
        sheet.SetRowHeight(3, points);
        sheet.DefaultRowHeight = points;

        var layout = SheetLayout.FromWorksheet(sheet, 7);

        Assert.Equal(pixels, layout.RowHeight(3));
        Assert.Equal(pixels, layout.RowHeight(4));
    }

    [Fact]
    public void Hidden_rows_and_columns_have_no_size()
    {
        var sheet = Sheet();
        sheet.SetColumnHidden(2, true);
        sheet.SetColumnWidth(2, 30);
        sheet.SetRowHidden(5, true);

        var layout = SheetLayout.FromWorksheet(sheet, 7);

        Assert.Equal(0, layout.ColumnWidth(2));
        Assert.Equal(0, layout.RowHeight(5));
    }

    [Fact]
    public void A_wider_font_widens_the_default_columns()
    {
        // 8 digits and 5 pixels of padding, rounded up to a multiple of 8 as Excel does.
        Assert.Equal(72, SheetLayout.FromWorksheet(Sheet(), maxDigitWidth: 8).ColumnWidth(1));
    }
}
