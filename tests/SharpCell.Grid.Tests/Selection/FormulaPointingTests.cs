namespace SharpCell.Grid.Tests.Selection;

/// <summary>Where a formula being typed takes a reference picked with the mouse, and how it is written.</summary>
public class FormulaPointingTests
{
    [Theory]
    [InlineData("=", 1, true)]
    [InlineData("=A1+", 4, true)]
    [InlineData("=SUM(", 5, true)]
    [InlineData("=SUM(A1, ", 9, true)]          // spaces after the comma
    [InlineData("=A1*(B2-", 8, true)]
    [InlineData("=A1>=", 5, true)]
    [InlineData("=A1", 3, false)]               // after a reference: a click enters the formula
    [InlineData("=SUM(A1)", 8, false)]
    [InlineData("=\"a,", 4, false)]             // inside text
    [InlineData("12+", 3, false)]               // not a formula
    [InlineData("=A1+", 2, false)]              // the caret inside the reference
    public void A_reference_goes_where_an_operand_is_expected(string text, int caret, bool expected)
    {
        Assert.Equal(expected, FormulaPointing.CanPoint(text, caret));
    }

    [Fact]
    public void A_reference_names_its_sheet_when_it_is_on_another()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        var other = workbook.AddSheet("Q1 Sales");

        Assert.Equal("B3", FormulaPointing.Reference(new CellArea(3, 2, 3, 2), sheet, sheet));
        Assert.Equal("B3:C5", FormulaPointing.Reference(new CellArea(3, 2, 5, 3), sheet, sheet));
        Assert.Equal("C:D", FormulaPointing.Reference(new CellArea(1, 3, SheetLayout.MaxRow, 4), sheet, sheet));
        Assert.Equal("2:2", FormulaPointing.Reference(new CellArea(2, 1, 2, SheetLayout.MaxColumn), sheet, sheet));
        Assert.Equal("'Q1 Sales'!A1", FormulaPointing.Reference(new CellArea(1, 1, 1, 1), other, sheet));
    }
}
