namespace SharpCell.Grid.Tests.Layout;

public class ColumnNamesTests
{
    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")]
    public void Column_letters(int column, string name)
    {
        Assert.Equal(name, ColumnNames.Of(column));
    }
}
