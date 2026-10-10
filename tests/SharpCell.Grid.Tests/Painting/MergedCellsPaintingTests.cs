namespace SharpCell.Grid.Tests.Painting;

public class MergedCellsPaintingTests
{
    private static readonly GridViewport Viewport = new(0, 0, 600, 300);

    private static Worksheet Sheet(string merge = "B2:D3")
    {
        var sheet = new Workbook().AddSheet("S");
        sheet["B2"].Value = "Title";
        sheet["C2"].Value = "hidden by the merge";
        sheet.Merge(merge);
        return sheet;
    }

    private static (RecordingCanvas Canvas, GridPainter Painter) Paint(Worksheet sheet, CellPosition? active = null, GridViewport? viewport = null)
    {
        var painter = new GridPainter { MeasureText = (text, _) => text.Length * 7 };
        var canvas = new RecordingCanvas();
        painter.Paint(canvas, sheet, viewport ?? Viewport, active);
        return (canvas, painter);
    }

    [Fact]
    public void The_top_left_cell_shows_across_the_whole_area()
    {
        var (canvas, painter) = Paint(Sheet());

        var text = canvas.TextOf("Title");
        var area = painter.CellBounds(Viewport, Cells.Parse("B2"), Cells.Parse("D3"));
        Assert.True(text.Bounds.X >= area.X && text.Bounds.Right <= area.Right && text.Bounds.Width > 3 * 64 - 10);
        Assert.DoesNotContain(canvas.Texts, t => t.Value == "hidden by the merge");
    }

    [Fact]
    public void No_gridline_crosses_a_merged_area()
    {
        var (canvas, painter) = Paint(Sheet());

        // Painted over the gridlines in the background colour, as Excel shows a merged area.
        var area = painter.CellBounds(Viewport, Cells.Parse("B2"), Cells.Parse("D3"));
        var cover = canvas.Calls.FindIndex(c => c is RecordingCanvas.Fill f && f.Rect.X == area.X && f.Rect.Y == area.Y && f.Rect.Width == area.Width - 1);
        var lastGridline = canvas.Calls.FindLastIndex(c => c is RecordingCanvas.Line l && l.Color == GridTheme.Excel.Gridline);
        Assert.True(cover > lastGridline, $"cover at {cover}, last gridline at {lastGridline}");
    }

    [Fact]
    public void A_merge_whose_top_left_cell_is_scrolled_away_still_shows()
    {
        var scrolled = Viewport with { OffsetX = 64 * 2 };   // column C first

        Assert.Single(Paint(Sheet(), viewport: scrolled).Canvas.Texts, t => t.Value == "Title");
    }

    [Fact]
    public void Clicking_inside_a_merge_hits_its_top_left_cell()
    {
        var painter = new GridPainter();
        var sheet = Sheet();
        var inside = painter.CellBounds(Viewport, Cells.Parse("D3"));

        Assert.Equal(Cells.Parse("B2"), painter.HitTest(Viewport, inside.X + 5, inside.Y + 5, sheet));
        Assert.Equal(Cells.Parse("E3"), painter.HitTest(Viewport, inside.Right + 5, inside.Y + 5, sheet));
    }

    [Fact]
    public void The_active_border_goes_round_the_whole_merge()
    {
        var (canvas, painter) = Paint(Sheet(), Cells.Parse("B2"));

        var border = Assert.Single(canvas.Calls.OfType<RecordingCanvas.Stroke>());
        var area = painter.CellBounds(Viewport, Cells.Parse("B2"), Cells.Parse("D3"));
        Assert.Equal(area.Width, border.Rect.Width);
        Assert.Equal(area.Height, border.Rect.Height);
    }

    [Theory]
    [InlineData("B2", 0, 1, "E2")]    // right out of the merge
    [InlineData("B2", 1, 0, "B4")]    // down out of the merge
    [InlineData("A2", 0, 1, "B2")]    // into it: its top-left cell
    [InlineData("C4", -1, 0, "B2")]   // up into it from below
    [InlineData("E3", 0, -1, "B2")]   // left into it
    [InlineData("B2", 0, -1, "A2")]
    public void Moving_steps_over_a_merge_as_one_cell(string from, int rows, int columns, string to)
    {
        var target = GridNavigation.Move(Sheet(), Cells.Parse(from), rows, columns);

        Assert.Equal(Cells.Parse(to), target);
    }
}
