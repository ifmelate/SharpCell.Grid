namespace SharpCell.Grid.Tests.Selection;

/// <summary>Ctrl+A selects the data around the active cell first, then the whole sheet, as in Excel.</summary>
public class SelectAllTests
{
    // B2:D4 filled, with C3 empty inside; E5 touches it diagonally; H1 stands alone; A10 too.
    private static Worksheet Sheet()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        foreach (var address in new[] { "B2", "C2", "D2", "B3", "D3", "B4", "C4", "D4", "E5", "H1", "A10" })
            sheet[address].Value = 1;
        return sheet;
    }

    [Theory]
    [InlineData("C3", "B2:E5")]     // an empty cell inside, and a diagonal neighbour joins
    [InlineData("B2", "B2:E5")]
    [InlineData("H1", "H1")]        // alone, its region is itself
    [InlineData("F6", "B2:F6")]     // empty, touching the region diagonally: it joins it
    [InlineData("J20", "J20")]      // empty and alone
    public void The_current_region_is_the_data_bounded_by_empty_rows_and_columns(string cell, string region)
    {
        var area = GridNavigation.CurrentRegion(Sheet(), Cells.Parse(cell));

        Assert.Equal(region, GridSelection.Cell(new CellPosition(area.FirstRow, area.FirstColumn), null).ExtendTo(new CellPosition(area.LastRow, area.LastColumn), null).ToString());
    }

    [Fact]
    public void Ctrl_A_selects_the_region_then_the_sheet()
    {
        var sheet = Sheet();
        var selection = GridSelection.Cell(Cells.Parse("C4"), sheet);

        var first = selection.SelectAll(sheet);
        var second = first.SelectAll(sheet);

        Assert.Equal(("B2:E5", Cells.Parse("C4")), (first.ToString(), first.ActiveCell));
        Assert.True(second.IsEntireRows && second.IsEntireColumns);
        Assert.Equal(Cells.Parse("C4"), second.ActiveCell);
    }

    [Fact]
    public void Ctrl_A_on_an_empty_cell_or_a_lone_one_selects_the_sheet()
    {
        var sheet = Sheet();

        Assert.True(GridSelection.Cell(Cells.Parse("J20"), sheet).SelectAll(sheet).IsEntireRows);
        Assert.True(GridSelection.Cell(Cells.Parse("H1"), sheet).SelectAll(sheet).IsEntireRows);
        Assert.True(GridSelection.Cell(Cells.Parse("C3"), null).SelectAll(null).IsEntireRows);
    }
}
