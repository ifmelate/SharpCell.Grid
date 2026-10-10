namespace SharpCell.Grid.Tests.Painting;

/// <summary>Text wider than its cell runs on over empty neighbours, as in Excel; 7 px per character.</summary>
public class TextOverflowTests
{
    private const string Long = "a long piece of text";   // 140 px, more than two default columns
    private static readonly GridViewport Viewport = new(0, 0, 600, 200);

    private static (Workbook Workbook, Worksheet Sheet) NewSheet()
    {
        var workbook = new Workbook();
        return (workbook, workbook.AddSheet("S"));
    }

    private static (RecordingCanvas Canvas, GridPainter Painter) Paint(Worksheet sheet, GridViewport? viewport = null)
    {
        var painter = new GridPainter { MeasureText = (text, _) => text.Length * 7 };
        var canvas = new RecordingCanvas();
        painter.Paint(canvas, sheet, viewport ?? Viewport);
        return (canvas, painter);
    }

    private static GridRect Bounds(GridPainter painter, string address, GridViewport? viewport = null) =>
        painter.CellBounds(viewport ?? Viewport, Cells.Parse(address));

    [Fact]
    public void Left_aligned_text_runs_on_to_the_right_over_empty_cells()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Value = Long;

        var (canvas, painter) = Paint(sheet);

        var text = canvas.TextOf(Long);
        Assert.True(text.Bounds.Width >= 140, $"{text.Bounds.Width} px");
        Assert.Equal(GridTextAlignment.Left, text.Alignment);
        Assert.True(text.Bounds.X < Bounds(painter, "C2").X && text.Bounds.Right <= Bounds(painter, "D2").Right);
    }

    [Fact]
    public void A_cell_with_content_stops_the_text()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Value = Long;
        sheet["D2"].Value = 1;

        var (canvas, painter) = Paint(sheet);

        Assert.True(canvas.TextOf(Long).Bounds.Right <= Bounds(painter, "C2").Right);
    }

    [Fact]
    public void A_formula_that_gives_empty_text_stops_the_text_too()
    {
        var (workbook, sheet) = NewSheet();
        sheet["B2"].Value = Long;
        sheet["C2"].Formula = "=\"\"";
        workbook.Recalculate();

        var (canvas, painter) = Paint(sheet);

        Assert.True(canvas.TextOf(Long).Bounds.Right <= Bounds(painter, "B2").Right);
    }

    [Fact]
    public void Right_aligned_text_runs_on_to_the_left()
    {
        var (_, sheet) = NewSheet();
        sheet["E2"].Value = Long;
        sheet["E2"].Style = CellStyle.Default with { HorizontalAlignment = CellHorizontalAlignment.Right };

        var (canvas, painter) = Paint(sheet);

        var text = canvas.TextOf(Long);
        Assert.True(text.Bounds.X < Bounds(painter, "D2").X && text.Bounds.Right <= Bounds(painter, "E2").Right);
    }

    [Fact]
    public void Centred_text_runs_on_to_both_sides()
    {
        var (_, sheet) = NewSheet();
        sheet["D2"].Value = Long;
        sheet["D2"].Style = CellStyle.Default with { HorizontalAlignment = CellHorizontalAlignment.Center };

        var (canvas, painter) = Paint(sheet);

        var text = canvas.TextOf(Long);
        Assert.True(text.Bounds.X < Bounds(painter, "D2").X && text.Bounds.Right > Bounds(painter, "D2").Right);
    }

    [Fact]
    public void Wrapped_text_numbers_and_merged_cells_stay_in_place()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Value = Long;
        sheet["B2"].Style = CellStyle.Default with { WrapText = true };
        sheet["B3"].Value = 123456789012345.0;
        sheet["B4"].Value = Long;
        sheet.Merge("B4:C4");

        var (canvas, painter) = Paint(sheet);

        Assert.True(canvas.Texts.First(t => t.Value == Long && t.Wrap).Bounds.Right <= Bounds(painter, "B2").Right);
        Assert.All(canvas.Texts.Where(t => t.Bounds.Y >= Bounds(painter, "B3").Y && t.Bounds.Bottom <= Bounds(painter, "B3").Bottom),
            t => Assert.True(t.Bounds.Right <= Bounds(painter, "B3").Right));
        Assert.True(canvas.Texts.Single(t => t.Value == Long && !t.Wrap).Bounds.Right <= Bounds(painter, "C4").Right);
    }

    [Fact]
    public void Text_stops_before_a_merged_area()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Value = Long;
        sheet.Merge("C2:D2");

        var (canvas, painter) = Paint(sheet);

        Assert.True(canvas.TextOf(Long).Bounds.Right <= Bounds(painter, "B2").Right);
    }

    [Fact]
    public void The_gridlines_under_running_text_are_covered()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Value = Long;

        var (canvas, painter) = Paint(sheet);

        // A one-pixel strip in the background colour over the gridline between B2 and C2.
        var edge = Bounds(painter, "B2").Right;
        var row = Bounds(painter, "B2");
        Assert.Contains(canvas.Calls.OfType<RecordingCanvas.Fill>(), f =>
            f.Rect.Width == 1 && f.Rect.X == Math.Floor(edge) - 1 && f.Rect.Y == row.Y && f.Color == GridTheme.Excel.Background);
    }

    [Fact]
    public void Text_from_a_cell_scrolled_out_of_view_still_shows()
    {
        var (_, sheet) = NewSheet();
        sheet["A2"].Value = Long;
        var scrolled = Viewport with { OffsetX = 64 };   // column B first

        var (canvas, painter) = Paint(sheet, scrolled);

        var text = canvas.TextOf(Long);
        Assert.True(text.Bounds.X < Bounds(painter, "B2", scrolled).X);
    }
}
