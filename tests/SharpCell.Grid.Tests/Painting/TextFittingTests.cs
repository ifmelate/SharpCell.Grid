namespace SharpCell.Grid.Tests.Painting;

/// <summary>What a cell shows when its text is wider than the column, with every character 7 px wide.</summary>
public class TextFittingTests
{
    private const double CharWidth = 7;
    private static readonly GridViewport Viewport = new(0, 0, 400, 200);

    private static GridPainter Painter() => new() { MeasureText = (text, _) => text.Length * CharWidth };

    private static string Shown(object value, string? format = null, string? formula = null, double columnWidth = SheetLayout.ExcelColumnWidth)
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        if (formula is not null)
            sheet["A1"].Formula = formula;
        else
            sheet["A1"].Value = value switch
            {
                string s => s,
                double d => d,
                int i => i,
                _ => throw new ArgumentException("Unsupported value", nameof(value)),
            };
        if (format is not null)
            sheet["A1"].NumberFormat = format;
        workbook.Recalculate();

        var painter = Painter();
        painter.Layout.SetColumnWidth(1, columnWidth);
        var canvas = new RecordingCanvas();
        painter.Paint(canvas, sheet, Viewport);
        var cell = painter.CellBounds(Viewport, new CellPosition(1, 1));
        return Assert.Single(canvas.Texts, t => t.Bounds.X >= cell.X && t.Bounds.X < cell.Right && t.Bounds.Y >= cell.Y && t.Bounds.Y < cell.Bottom).Value;
    }

    // A 64 px column leaves 59 px for text: 8 characters of 7 px.

    [Fact]
    public void A_number_that_fits_is_unchanged() => Assert.Equal("1,250.00", Shown(1250, "#,##0.00"));

    [Fact]
    public void A_formatted_number_too_wide_shows_hashes() => Assert.Equal("########", Shown(45941, "yyyy-mm-dd"));

    [Fact]
    public void A_General_number_too_wide_loses_decimals() => Assert.Equal("0.333333", Shown(0, formula: "=1/3"));

    [Fact]
    public void A_General_number_rounds_when_it_loses_decimals() => Assert.Equal("0.666667", Shown(0, formula: "=2/3"));

    [Fact]
    public void A_large_General_number_keeps_fewer_digits_in_scientific_notation() =>
        Assert.Equal("1.23E+11", Shown(123456789012.0));

    [Fact]
    public void A_General_number_that_cannot_be_shortened_enough_shows_hashes() =>
        Assert.Equal("###", Shown(123456789012.0, columnWidth: 26));

    [Fact]
    public void Text_too_wide_is_cut_off_not_replaced() =>
        Assert.Equal("a long piece of text", Shown("a long piece of text"));

    [Fact]
    public void Row_headers_widen_for_long_row_numbers()
    {
        var painter = Painter();
        var bottom = Viewport with { OffsetY = painter.Layout.Height };

        var top = painter.CellBounds(Viewport, new CellPosition(1, 1));
        var end = painter.CellBounds(bottom, new CellPosition(SheetLayout.MaxRow, 1));

        Assert.Equal(painter.RowHeaderWidth, top.X);
        Assert.Equal(7 * CharWidth + 12, end.X);
    }
}
