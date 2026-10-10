namespace SharpCell.Grid.Tests;

/// <summary>Cell positions by A1 address, for readable tests.</summary>
internal static class Cells
{
    public static CellPosition Parse(string address)
    {
        var letters = address.TakeWhile(char.IsLetter).ToArray();
        var column = letters.Aggregate(0, (sum, letter) => sum * 26 + letter - 'A' + 1);
        return new CellPosition(int.Parse(address.AsSpan(letters.Length)), column);
    }
}
