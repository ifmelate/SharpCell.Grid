namespace SharpCell.Grid.Tests.Painting;

/// <summary>Frozen rows and columns stay in place while the rest scrolls, as Excel's Freeze Panes.</summary>
public class FrozenPaneTests
{
    // Scrolled ten columns right and ten rows down past the frozen row and column.
    private static readonly GridViewport Scrolled = new(64 * 10, 20 * 10, 400, 200);

    private static GridPainter Painter() => new() { FrozenRows = 1, FrozenColumns = 1 };

    private static RecordingCanvas Paint(Worksheet? sheet, GridPainter painter, GridViewport viewport)
    {
        var canvas = new RecordingCanvas();
        painter.Paint(canvas, sheet, viewport);
        return canvas;
    }

    [Fact]
    public void Frozen_cells_do_not_move_and_the_rest_scrolls_after_them()
    {
        var painter = Painter();

        var a1 = painter.CellBounds(Scrolled, Cells.Parse("A1"));
        var l12 = painter.CellBounds(Scrolled, Cells.Parse("L12"));

        Assert.Equal((painter.RowHeaderWidth, painter.ColumnHeaderHeight), (a1.X, a1.Y));
        Assert.Equal((a1.Right, a1.Bottom), (l12.X, l12.Y));   // the first scrolled-in cell sits right after them
    }

    [Fact]
    public void The_visible_area_is_the_scrolled_part()
    {
        Assert.Equal(12, Painter().VisibleArea(Scrolled).FirstColumn);
        Assert.Equal(12, Painter().VisibleArea(Scrolled).FirstRow);
    }

    [Fact]
    public void Hit_testing_tells_frozen_from_scrolled_cells()
    {
        var painter = Painter();
        var x = painter.RowHeaderWidth;
        var y = painter.ColumnHeaderHeight;

        Assert.Equal(Cells.Parse("A1"), painter.HitTest(Scrolled, x + 5, y + 5));
        Assert.Equal(Cells.Parse("L12"), painter.HitTest(Scrolled, x + 64 + 5, y + 20 + 5));
        Assert.Equal(Cells.Parse("A12"), painter.HitTest(Scrolled, x + 5, y + 20 + 5));
    }

    [Fact]
    public void Frozen_and_scrolled_cells_and_headers_are_painted_and_the_cells_between_are_not()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet["A1"].Value = "corner";
        sheet["L1"].Value = "top";
        sheet["A12"].Value = "left";
        sheet["L12"].Value = "main";
        sheet["B2"].Value = "scrolled away";

        var canvas = Paint(sheet, Painter(), Scrolled);

        foreach (var text in new[] { "corner", "top", "left", "main", "A", "L", "1", "12" })
            canvas.TextOf(text);
        Assert.DoesNotContain(canvas.Texts, t => t.Value is "scrolled away" or "B" or "2");
    }

    [Fact]
    public void A_line_marks_where_the_panes_are_frozen()
    {
        var painter = Painter();

        var lines = Paint(null, painter, Scrolled).Calls.OfType<RecordingCanvas.Line>().Where(l => l.Color == painter.Theme.FrozenLine).ToList();

        Assert.Contains(lines, l => l.X1 == l.X2 && Math.Abs(l.X1 - (painter.RowHeaderWidth + 64)) <= 1);
        Assert.Contains(lines, l => l.Y1 == l.Y2 && Math.Abs(l.Y1 - (painter.ColumnHeaderHeight + 20)) <= 1);
    }

    [Fact]
    public void Scrolling_into_view_leaves_frozen_cells_alone_and_scrolls_after_them()
    {
        var painter = Painter();
        var top = new GridViewport(0, 0, 400, 200);

        Assert.Equal(Scrolled, painter.ScrollIntoView(Scrolled, Cells.Parse("A1")));
        var down = painter.ScrollIntoView(top, Cells.Parse("A30"));
        Assert.Equal(top.Height, painter.CellBounds(down, Cells.Parse("A30")).Bottom);
        var back = painter.ScrollIntoView(Scrolled, Cells.Parse("B2"));
        Assert.Equal((0.0, 0.0), (back.OffsetX, back.OffsetY));
    }

    [Fact]
    public void Text_in_a_frozen_column_does_not_run_on_into_the_scrolled_part()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet["A15"].Value = "a long piece of text";
        var painter = Painter();
        painter.MeasureText = (text, _) => text.Length * 7;

        var text = Paint(sheet, painter, Scrolled).TextOf("a long piece of text");

        Assert.True(text.Bounds.Right <= painter.CellBounds(Scrolled, Cells.Parse("A15")).Right);
    }

    [Fact]
    public void A_sheet_without_gridlines_paints_none()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.ShowGridlines = false;
        var painter = new GridPainter();

        var canvas = Paint(sheet, painter, new GridViewport(0, 0, 400, 200));

        Assert.DoesNotContain(canvas.Calls.OfType<RecordingCanvas.Line>(), l => l.Color == painter.Theme.Gridline);
    }

    [Fact]
    public void Frozen_counts_cannot_be_negative()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new GridPainter { FrozenRows = -1 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new GridPainter { FrozenColumns = -1 });
    }
}
