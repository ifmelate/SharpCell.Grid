namespace SharpCell.Grid.Tests.Painting;

public class StylePaintingTests
{
    private static readonly GridViewport Viewport = new(0, 0, 600, 300);

    private static (Workbook Workbook, Worksheet Sheet) NewSheet()
    {
        var workbook = new Workbook();
        return (workbook, workbook.AddSheet("S"));
    }

    private static (RecordingCanvas Canvas, GridPainter Painter) Paint(Worksheet sheet, CellPosition? active = null)
    {
        var painter = new GridPainter { MeasureText = (text, _) => text.Length * 7 };
        painter.Layout = SheetLayout.FromWorksheet(sheet, 7);
        var canvas = new RecordingCanvas();
        painter.Paint(canvas, sheet, Viewport, active);
        return (canvas, painter);
    }

    private static CellPosition At(string address) => Cells.Parse(address);

    [Fact]
    public void A_fill_covers_the_whole_cell()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Style = CellStyle.Default with { Fill = CellColor.FromRgb(0xFFFF00) };

        var (canvas, painter) = Paint(sheet);

        var fill = Assert.Single(canvas.Calls.OfType<RecordingCanvas.Fill>(), f => f.Color == new GridColor(0xFF, 0xFF, 0x00));
        Assert.Equal(painter.CellBounds(Viewport, At("B2")), fill.Rect);
    }

    [Fact]
    public void A_fill_hides_the_gridlines_of_its_cell()
    {
        // Gridlines are painted first, so the fill lies over them.
        var (_, sheet) = NewSheet();
        sheet["B2"].Style = CellStyle.Default with { Fill = CellColor.FromRgb(0xFFFF00) };

        var (canvas, _) = Paint(sheet);

        var lastGridline = canvas.Calls.FindLastIndex(c => c is RecordingCanvas.Line l && l.Color == GridTheme.Excel.Gridline);
        var fill = canvas.Calls.FindIndex(c => c is RecordingCanvas.Fill f && f.Color == new GridColor(0xFF, 0xFF, 0x00));
        Assert.True(fill > lastGridline);
    }

    [Fact]
    public void Text_is_drawn_in_the_cells_font_and_colour()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = "Total";
        sheet["A1"].Style = CellStyle.Default with
        {
            Font = new CellFont { Name = "Arial", Size = 12, Bold = true, Italic = true, Underline = CellUnderline.Single, Strikethrough = true, Color = CellColor.FromRgb(0xC00000) },
        };

        var text = Paint(sheet).Canvas.TextOf("Total");

        Assert.Equal(new GridFont("Arial", 16, Bold: true, Italic: true, Underline: true, Strikethrough: true), text.Font);
        Assert.Equal(new GridColor(0xC0, 0x00, 0x00), text.Color);
    }

    [Fact]
    public void Cells_without_a_style_use_the_workbook_font_at_96_dpi()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = "x";

        var text = Paint(sheet).Canvas.TextOf("x");

        Assert.Equal(new GridFont("Calibri", 11 * 96 / 72.0), text.Font);
        Assert.Equal(GridTheme.Excel.CellText, text.Color);
    }

    [Theory]
    [InlineData(CellHorizontalAlignment.Left, GridTextAlignment.Left)]
    [InlineData(CellHorizontalAlignment.Center, GridTextAlignment.Center)]
    [InlineData(CellHorizontalAlignment.Right, GridTextAlignment.Right)]
    [InlineData(CellHorizontalAlignment.CenterContinuous, GridTextAlignment.Center)]
    [InlineData(CellHorizontalAlignment.Distributed, GridTextAlignment.Center)]
    [InlineData(CellHorizontalAlignment.Justify, GridTextAlignment.Left)]
    [InlineData(CellHorizontalAlignment.Fill, GridTextAlignment.Left)]
    public void Horizontal_alignment_overrides_General(CellHorizontalAlignment set, GridTextAlignment painted)
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = 42;
        sheet["A1"].Style = CellStyle.Default with { HorizontalAlignment = set };

        Assert.Equal(painted, Paint(sheet).Canvas.TextOf("42").Alignment);
    }

    [Theory]
    [InlineData(CellVerticalAlignment.Top, GridVerticalAlignment.Top)]
    [InlineData(CellVerticalAlignment.Center, GridVerticalAlignment.Center)]
    [InlineData(CellVerticalAlignment.Bottom, GridVerticalAlignment.Bottom)]
    [InlineData(CellVerticalAlignment.Justify, GridVerticalAlignment.Center)]
    public void Vertical_alignment_is_followed(CellVerticalAlignment set, GridVerticalAlignment painted)
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = "x";
        sheet["A1"].Style = CellStyle.Default with { VerticalAlignment = set };

        Assert.Equal(painted, Paint(sheet).Canvas.TextOf("x").VerticalAlignment);
    }

    [Fact]
    public void Wrapped_text_is_drawn_wrapped_and_never_shortened()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = "a long piece of text";
        sheet["A1"].Style = CellStyle.Default with { WrapText = true };

        Assert.True(Paint(sheet).Canvas.TextOf("a long piece of text").Wrap);
    }

    [Fact]
    public void An_indent_moves_left_aligned_text_right()
    {
        var (_, sheet) = NewSheet();
        sheet["A1"].Value = "x";
        sheet["A2"].Value = "x2";
        sheet["A2"].Style = CellStyle.Default with { HorizontalAlignment = CellHorizontalAlignment.Left, Indent = 2 };

        var (canvas, _) = Paint(sheet);

        Assert.Equal(canvas.TextOf("x").Bounds.X + 2 * 9, canvas.TextOf("x2").Bounds.X);
    }

    [Theory]
    [InlineData(CellBorderStyle.Thin, 1, GridDash.Solid, 1)]
    [InlineData(CellBorderStyle.Medium, 2, GridDash.Solid, 1)]
    [InlineData(CellBorderStyle.Thick, 3, GridDash.Solid, 1)]
    [InlineData(CellBorderStyle.Double, 1, GridDash.Solid, 2)]
    [InlineData(CellBorderStyle.Dotted, 1, GridDash.Dot, 1)]
    [InlineData(CellBorderStyle.Hair, 1, GridDash.Hair, 1)]
    [InlineData(CellBorderStyle.Dashed, 1, GridDash.Dash, 1)]
    [InlineData(CellBorderStyle.MediumDashed, 2, GridDash.Dash, 1)]
    [InlineData(CellBorderStyle.DashDot, 1, GridDash.DashDot, 1)]
    [InlineData(CellBorderStyle.MediumDashDotDot, 2, GridDash.DashDotDot, 1)]
    public void Border_lines_follow_their_style(CellBorderStyle style, double thickness, GridDash dash, int lines)
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Style = CellStyle.Default with { BottomBorder = new CellBorder(style, CellColor.FromRgb(0x0000FF)) };

        var (canvas, painter) = Paint(sheet);

        var drawn = canvas.Calls.OfType<RecordingCanvas.Line>().Where(l => l.Color == new GridColor(0, 0, 0xFF)).ToList();
        Assert.Equal(lines, drawn.Count);
        Assert.All(drawn, l => Assert.Equal((thickness, dash), (l.Thickness, l.Dash)));
        var cell = painter.CellBounds(Viewport, At("B2"));
        Assert.All(drawn, l => Assert.True(l.Y1 == l.Y2 && Math.Abs(l.Y1 - cell.Bottom) <= 2, $"line at {l.Y1}, cell ends at {cell.Bottom}"));
    }

    [Fact]
    public void A_border_without_a_colour_is_black()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Style = CellStyle.Default with { LeftBorder = new CellBorder(CellBorderStyle.Thin) };

        Assert.Single(Paint(sheet).Canvas.Calls.OfType<RecordingCanvas.Line>(), l => l.Color == new GridColor(0, 0, 0));
    }

    [Fact]
    public void Columns_and_rows_take_their_sizes_from_the_layout()
    {
        var (_, sheet) = NewSheet();
        sheet.SetColumnWidth(1, 21.42578125);
        sheet.SetRowHeight(1, 30);

        var (_, painter) = Paint(sheet);

        var cell = painter.CellBounds(Viewport, At("A1"));
        Assert.Equal((150.0, 40.0), (cell.Width, cell.Height));
    }
}
