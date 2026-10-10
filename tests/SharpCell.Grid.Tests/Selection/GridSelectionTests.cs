namespace SharpCell.Grid.Tests.Selection;

public class GridSelectionTests
{
    private static Worksheet Sheet()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.Merge("C3:D4");
        return sheet;
    }

    private static GridSelection Cell(string address, Worksheet? sheet = null) => GridSelection.Cell(Cells.Parse(address), sheet);

    [Fact]
    public void One_cell_is_selected_alone()
    {
        var selection = Cell("B2");

        Assert.Equal("B2", selection.ToString());
        Assert.True(selection.IsSingleCell);
        Assert.Equal(Cells.Parse("B2"), selection.ActiveCell);
    }

    [Fact]
    public void A_cell_in_a_merged_area_selects_the_area_as_one_cell()
    {
        var selection = Cell("C3", Sheet());

        Assert.Equal("C3:D4", selection.ToString());
        Assert.True(selection.IsSingleCell);
    }

    [Theory]
    [InlineData("B2", 0, 1, "B2:C2")]
    [InlineData("B2", 1, 1, "B2:C3")]
    [InlineData("B2", -1, -1, "A1:B2")]
    [InlineData("A1", -1, 0, "A1")]   // against the sheet's edge nothing changes
    public void Shift_with_the_arrows_moves_the_far_corner(string from, int rows, int columns, string expected)
    {
        var selection = Cell(from).Extend(rows, columns, null);

        Assert.Equal(expected, selection.ToString());
        Assert.Equal(Cells.Parse(from), selection.ActiveCell);
    }

    [Fact]
    public void Extending_back_shrinks_the_selection()
    {
        var selection = Cell("B2").Extend(0, 2, null).Extend(0, -1, null);

        Assert.Equal("B2:C2", selection.ToString());
    }

    [Fact]
    public void Extending_into_a_merged_area_takes_it_whole()
    {
        var selection = Cell("B3", Sheet()).Extend(0, 1, Sheet());

        Assert.Equal("B3:D4", selection.ToString());
    }

    [Fact]
    public void A_step_a_merged_area_swallows_goes_on_until_the_selection_changes()
    {
        var sheet = Sheet();

        // B3:C3 already covers C3:D4; the next step right reaches E.
        var selection = Cell("B3", sheet).Extend(0, 1, sheet).Extend(0, 1, sheet);

        Assert.Equal("B3:E4", selection.ToString());
    }

    [Fact]
    public void An_arrow_without_shift_selects_the_next_cell_alone()
    {
        var selection = Cell("B2").Extend(2, 2, null).Move(0, 1, null);

        Assert.Equal("C2", selection.ToString());
        Assert.True(selection.IsSingleCell);
    }

    [Fact]
    public void Enter_goes_down_each_column_of_a_selection_and_wraps_around()
    {
        var selection = Cell("B2").ExtendTo(Cells.Parse("C3"), null);

        var visited = new List<string>();
        for (var i = 0; i < 5; i++)
        {
            selection = selection.MoveWithin(1, 0, null);
            visited.Add(selection.ActiveCell.ToString());
        }

        Assert.Equal(["B3", "C2", "C3", "B2", "B3"], visited);
        Assert.Equal("B2:C3", selection.ToString());
    }

    [Fact]
    public void Tab_goes_across_each_row_and_shift_tab_back()
    {
        var selection = Cell("B2").ExtendTo(Cells.Parse("C3"), null);

        Assert.Equal("C2", selection.MoveWithin(0, 1, null).ActiveCell.ToString());
        Assert.Equal("B3", selection.MoveWithin(0, 1, null).MoveWithin(0, 1, null).ActiveCell.ToString());
        Assert.Equal("C3", selection.MoveWithin(0, -1, null).ActiveCell.ToString());
    }

    [Fact]
    public void Enter_passes_over_hidden_rows_and_the_inner_cells_of_merged_areas()
    {
        var sheet = Sheet();
        sheet.SetRowHidden(2, true);
        var selection = Cell("C1", sheet).ExtendTo(Cells.Parse("C5"), sheet);

        var visited = new List<string>();
        for (var i = 0; i < 4; i++)
        {
            selection = selection.MoveWithin(1, 0, sheet);
            visited.Add(selection.ActiveCell.ToString());
        }

        // C1:C5 covers C3:D4, so it is C1:D5: C2 hidden, C4, D3 and D4 inside the merge.
        Assert.Equal(["C3", "C5", "D1", "D5"], visited);
    }

    [Fact]
    public void Enter_on_a_single_cell_moves_as_the_arrow_does()
    {
        Assert.Equal("B3", Cell("B2").MoveWithin(1, 0, null).ToString());
    }

    [Fact]
    public void Whole_columns_and_rows_have_their_own_addresses()
    {
        Assert.Equal("C:E", GridSelection.Columns(3, 5, 10, null).ToString());
        Assert.Equal("C:E", GridSelection.Columns(5, 3, 10, null).ToString());
        Assert.Equal("2:4", GridSelection.Rows(4, 2, 1, null).ToString());
        Assert.Equal("1:1048576", GridSelection.All(Cells.Parse("B2")).ToString());
        Assert.Equal(Cells.Parse("C10"), GridSelection.Columns(3, 5, 10, null).ActiveCell);
    }

    [Fact]
    public void Extending_whole_columns_keeps_them_whole()
    {
        var selection = GridSelection.Columns(3, 3, 7, null).ExtendTo(Cells.Parse("E20"), null);

        Assert.Equal("C:E", selection.ToString());
        Assert.True(selection.IsEntireColumns);
        Assert.Equal(Cells.Parse("C7"), selection.ActiveCell);
    }

    [Fact]
    public void Shift_clicking_a_column_header_selects_whole_columns_from_the_anchor()
    {
        var selection = Cell("B5").ExtendTo(new GridHit(GridPart.ColumnHeader, Cells.Parse("D1")), null);

        Assert.Equal("B:D", selection.ToString());
        Assert.Equal(Cells.Parse("B5"), selection.ActiveCell);
    }

    [Fact]
    public void Ctrl_space_and_shift_space_take_the_whole_columns_and_rows_of_a_selection()
    {
        var selection = Cell("B2").ExtendTo(Cells.Parse("C4"), null);

        Assert.Equal("B:C", selection.EntireColumns(null).ToString());
        Assert.Equal("2:4", selection.EntireRows(null).ToString());
        Assert.Equal(Cells.Parse("B2"), selection.EntireRows(null).ActiveCell);
    }

    [Theory]
    [InlineData(GridPart.Cell, "B2")]
    [InlineData(GridPart.ColumnHeader, "B:B")]
    [InlineData(GridPart.RowHeader, "2:2")]
    [InlineData(GridPart.Corner, "1:1048576")]
    public void A_click_selects_what_it_hits(GridPart part, string expected)
    {
        var selection = GridSelection.At(new GridHit(part, Cells.Parse("B2")), null);

        Assert.Equal(expected, selection.ToString());
        Assert.Equal(Cells.Parse("B2"), selection.ActiveCell);
    }

    [Fact]
    public void Selections_compare_by_value()
    {
        Assert.Equal(Cell("B2").Extend(1, 1, null), Cell("B2").ExtendTo(Cells.Parse("C3"), null));
        Assert.NotEqual(Cell("B2"), Cell("B3"));
    }
}
