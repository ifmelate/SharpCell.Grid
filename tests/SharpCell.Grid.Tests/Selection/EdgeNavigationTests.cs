namespace SharpCell.Grid.Tests.Selection;

public class EdgeNavigationTests
{
    // A1:A3 filled, A4 empty, A7:A8 filled; C1 filled; B5 has a formula giving "".
    private static Worksheet Sheet()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        foreach (var address in new[] { "A1", "A2", "A3", "A7", "A8", "C1" })
            sheet[address].Value = 1;
        sheet["B5"].Formula = "=\"\"";
        workbook.Recalculate();
        return sheet;
    }

    [Theory]
    [InlineData("A1", 1, 0, "A3")]        // to the end of the block
    [InlineData("A3", 1, 0, "A7")]        // over the gap to the next block
    [InlineData("A8", 1, 0, "A1048576")]  // nothing below: the sheet's edge
    [InlineData("A7", -1, 0, "A3")]
    [InlineData("A1", 0, 1, "C1")]
    [InlineData("C1", 0, 1, "XFD1")]
    [InlineData("B1", 1, 0, "B5")]        // a formula is content, even one that shows nothing
    [InlineData("A1", -1, 0, "A1")]       // at the edge already
    public void Ctrl_arrows_jump_as_in_Excel(string from, int rows, int columns, string to)
    {
        Assert.Equal(Cells.Parse(to), GridNavigation.MoveToEdge(Sheet(), Cells.Parse(from), rows, columns));
    }

    [Fact]
    public void Ctrl_shift_arrows_extend_to_the_edge()
    {
        var sheet = Sheet();

        var selection = GridSelection.Cell(Cells.Parse("A1"), sheet).ExtendToEdge(1, 0, sheet).ExtendToEdge(1, 0, sheet);

        Assert.Equal("A1:A7", selection.ToString());
        Assert.Equal(Cells.Parse("A1"), selection.ActiveCell);
    }
}
