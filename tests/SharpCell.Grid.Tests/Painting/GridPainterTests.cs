namespace SharpCell.Grid.Tests.Painting;

public class GridPainterTests
{
    private static readonly GridViewport Small = new(0, 0, 400, 200);

    private static (Workbook Workbook, Worksheet Sheet) NewSheet()
    {
        var workbook = new Workbook();
        return (workbook, workbook.AddSheet("Sheet1"));
    }

    private static RecordingCanvas Paint(Worksheet? sheet, GridViewport viewport, CellPosition? active = null, GridPainter? painter = null)
    {
        var canvas = new RecordingCanvas();
        (painter ?? new GridPainter()).Paint(canvas, sheet, viewport, active);
        return canvas;
    }

    [Fact]
    public void Cells_show_their_text_as_Excel_formats_it()
    {
        var (workbook, sheet) = NewSheet();
        sheet["B2"].Value = 1.5;
        sheet["B2"].NumberFormat = "0.00";
        sheet["C2"].Formula = "=B2/12";
        sheet["C2"].NumberFormat = "0.0%";
        workbook.Recalculate();

        var canvas = Paint(sheet, Small);

        canvas.TextOf("1.50");
        canvas.TextOf("12.5%");
    }

    [Fact]
    public void Numbers_go_right_text_left_logical_values_and_errors_in_the_middle()
    {
        // Excel's General alignment.
        var (workbook, sheet) = NewSheet();
        sheet["A1"].Value = "Item";
        sheet["B1"].Value = 42;
        sheet["C1"].Value = true;
        sheet["D1"].Formula = "=1/0";
        workbook.Recalculate();

        var canvas = Paint(sheet, Small);

        Assert.Equal(GridTextAlignment.Left, canvas.TextOf("Item").Alignment);
        Assert.Equal(GridTextAlignment.Right, canvas.TextOf("42").Alignment);
        Assert.Equal(GridTextAlignment.Center, canvas.TextOf("TRUE").Alignment);
        Assert.Equal(GridTextAlignment.Center, canvas.TextOf("#DIV/0!").Alignment);
        Assert.All(canvas.Texts.Where(t => t.Value is "Item" or "42"), t => Assert.Equal(GridVerticalAlignment.Bottom, t.VerticalAlignment));
    }

    [Fact]
    public void Text_stays_inside_its_cell()
    {
        var (_, sheet) = NewSheet();
        sheet["B2"].Value = "x";
        var painter = new GridPainter();

        var text = Paint(sheet, Small, painter: painter).TextOf("x");

        var cell = painter.CellBounds(Small, new CellPosition(2, 2));
        Assert.True(text.Bounds.X >= cell.X && text.Bounds.Right <= cell.Right);
        Assert.True(text.Bounds.Y >= cell.Y && text.Bounds.Bottom <= cell.Bottom);
    }

    [Fact]
    public void Headers_name_the_visible_columns_and_rows()
    {
        var canvas = Paint(null, Small);

        // 400 px wide minus the row headers shows columns A to F; 200 px high minus the column headers rows 1 to 9.
        foreach (var name in new[] { "A", "B", "C", "D", "E", "F", "1", "9" })
            Assert.Equal(GridTextAlignment.Center, canvas.TextOf(name).Alignment);
        Assert.DoesNotContain(canvas.Texts, t => t.Value is "G" or "10");
    }

    [Fact]
    public void A_cell_is_placed_after_the_headers_and_moves_with_the_offset()
    {
        var painter = new GridPainter();

        var unscrolled = painter.CellBounds(Small, new CellPosition(1, 1));
        var scrolled = painter.CellBounds(Small with { OffsetX = 64, OffsetY = 20 }, new CellPosition(2, 2));

        Assert.Equal(new GridRect(painter.RowHeaderWidth, painter.ColumnHeaderHeight, 64, 20), unscrolled);
        Assert.Equal(unscrolled, scrolled);
    }

    [Fact]
    public void Only_the_visible_cells_are_painted_at_the_end_of_the_sheet()
    {
        var (_, sheet) = NewSheet();
        sheet[1, 1].Value = "top";
        sheet[SheetLayout.MaxRow, 1].Value = "bottom";
        var painter = new GridPainter();
        var bottom = new GridViewport(0, painter.Layout.Height - 180, 400, 200);

        var canvas = Paint(sheet, bottom, painter: painter);

        canvas.TextOf("bottom");
        canvas.TextOf("1048576");
        Assert.DoesNotContain(canvas.Texts, t => t.Value == "top");
        Assert.True(canvas.Calls.Count < 100, $"{canvas.Calls.Count} draw calls");
    }

    [Fact]
    public void The_active_cell_gets_a_border()
    {
        var painter = new GridPainter();

        var canvas = Paint(null, Small, new CellPosition(3, 2), painter);

        Assert.Equal(painter.CellBounds(Small, new CellPosition(3, 2)), canvas.Outline(painter));
    }

    [Fact]
    public void An_active_cell_out_of_view_is_not_painted()
    {
        var painter = new GridPainter();

        var canvas = Paint(null, Small, new CellPosition(500, 2), painter);

        Assert.Null(canvas.Outline(painter));
    }

    [Fact]
    public void Hit_testing_finds_cells_and_ignores_headers()
    {
        var painter = new GridPainter();
        var scrolled = Small with { OffsetX = 64 * 10, OffsetY = 20 * 100 };

        Assert.Null(painter.HitTest(Small, 5, 100));
        Assert.Null(painter.HitTest(Small, 100, 5));
        Assert.Equal(new CellPosition(1, 1), painter.HitTest(Small, painter.RowHeaderWidth + 1, painter.ColumnHeaderHeight + 1));
        Assert.Equal(new CellPosition(101, 11), painter.HitTest(scrolled, painter.RowHeaderWidth + 1, painter.ColumnHeaderHeight + 1));
        Assert.Null(painter.HitTest(Small, 401, 100));
    }

    [Fact]
    public void Scrolling_a_cell_into_view_moves_as_little_as_possible()
    {
        var painter = new GridPainter();

        var right = painter.ScrollIntoView(Small, new CellPosition(1, 10));
        var below = painter.ScrollIntoView(Small, new CellPosition(30, 1));
        var back = painter.ScrollIntoView(right, new CellPosition(1, 1));
        var visible = painter.ScrollIntoView(Small, new CellPosition(2, 2));

        Assert.Equal(painter.CellBounds(right, new CellPosition(1, 10)).Right, Small.Width);
        Assert.Equal(painter.CellBounds(below, new CellPosition(30, 1)).Bottom, Small.Height);
        Assert.Equal(0, back.OffsetX);
        Assert.Equal(Small, visible);
    }

    [Fact]
    public void Offsets_are_kept_inside_the_sheet()
    {
        var painter = new GridPainter();

        var clamped = painter.Clamp(Small with { OffsetX = -10, OffsetY = 1e12 });

        Assert.Equal(0, clamped.OffsetX);
        Assert.Equal(painter.Layout.Height - (Small.Height - painter.ColumnHeaderHeight), clamped.OffsetY);
    }

    [Fact]
    public void Cell_positions_read_like_addresses()
    {
        Assert.Equal("B3", new CellPosition(3, 2).ToString());
        Assert.Throws<ArgumentOutOfRangeException>(() => new CellPosition(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new CellPosition(1, SheetLayout.MaxColumn + 1));
    }
}
