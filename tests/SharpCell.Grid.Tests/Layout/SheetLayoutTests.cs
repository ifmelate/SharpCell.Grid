namespace SharpCell.Grid.Tests.Layout;

public class SheetLayoutTests
{
    [Fact]
    public void Default_sizes_are_Excels_at_100_percent()
    {
        // Calibri 11, the default of a new workbook: columns 8.43 characters (64 px), rows 15 pt (20 px).
        var layout = new SheetLayout();

        Assert.Equal(64, layout.ColumnWidth(1));
        Assert.Equal(20, layout.RowHeight(1));
        Assert.Equal(64 * 3, layout.ColumnLeft(4));
        Assert.Equal(20 * 9, layout.RowTop(10));
    }

    [Fact]
    public void Whole_sheet_has_Excels_row_and_column_count()
    {
        var layout = new SheetLayout();

        Assert.Equal(16384 * 64.0, layout.Width);
        Assert.Equal(1048576 * 20.0, layout.Height);
    }

    [Fact]
    public void A_changed_width_moves_the_columns_after_it()
    {
        var layout = new SheetLayout();
        layout.SetColumnWidth(2, 100);
        layout.SetColumnWidth(5, 10);

        Assert.Equal(100, layout.ColumnWidth(2));
        Assert.Equal(64, layout.ColumnLeft(2));
        Assert.Equal(164, layout.ColumnLeft(3));
        Assert.Equal(164 + 64 + 64, layout.ColumnLeft(5));
        Assert.Equal(164 + 64 + 64 + 10, layout.ColumnLeft(6));
        Assert.Equal(16384 * 64.0 + 36 - 54, layout.Width);
    }

    [Fact]
    public void Setting_the_default_size_back_forgets_the_change()
    {
        var layout = new SheetLayout();
        layout.SetRowHeight(3, 40);
        layout.SetRowHeight(3, 20);

        Assert.Equal(20 * 9, layout.RowTop(10));
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(63.9, 1)]
    [InlineData(64, 2)]
    [InlineData(-5, 1)]
    [InlineData(1e12, 16384)]
    public void Column_at_an_offset(double x, int column)
    {
        Assert.Equal(column, new SheetLayout().ColumnAt(x));
    }

    [Fact]
    public void Hit_testing_follows_changed_sizes()
    {
        var layout = new SheetLayout();
        layout.SetRowHeight(2, 100);

        Assert.Equal(1, layout.RowAt(19));
        Assert.Equal(2, layout.RowAt(20));
        Assert.Equal(2, layout.RowAt(119));
        Assert.Equal(3, layout.RowAt(120));
        Assert.Equal(1000, layout.RowAt(layout.RowTop(1000) + 1));
    }

    [Fact]
    public void A_row_of_height_zero_is_never_hit()
    {
        // Hidden rows have height zero; the offset where they would be belongs to the next row.
        var layout = new SheetLayout();
        layout.SetRowHeight(2, 0);

        Assert.Equal(20, layout.RowTop(2));
        Assert.Equal(20, layout.RowTop(3));
        Assert.Equal(3, layout.RowAt(20));
    }

    [Fact]
    public void Visible_cells_of_a_viewport()
    {
        var layout = new SheetLayout();

        var area = layout.VisibleArea(x: 100, y: 30, width: 200, height: 45);

        // Columns 2 (64..128) to 5 (256..320), rows 2 (20..40) to 4 (60..80).
        Assert.Equal(new CellArea(2, 2, 4, 5), area);
    }

    [Fact]
    public void Sizes_outside_the_sheet_are_refused()
    {
        var layout = new SheetLayout();

        Assert.Throws<ArgumentOutOfRangeException>(() => layout.SetColumnWidth(0, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => layout.SetColumnWidth(16385, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => layout.SetRowHeight(1, -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => layout.RowHeight(1048577));
    }
}
