namespace SharpCell.Grid.Tests.Painting;

public class SelectionPaintingTests
{
    private static readonly GridViewport Viewport = new(0, 0, 600, 300);

    private static (RecordingCanvas Canvas, GridPainter Painter) Paint(GridSelection? selection, Worksheet? sheet = null, GridPainter? painter = null, GridViewport? viewport = null)
    {
        painter ??= new GridPainter();
        var canvas = new RecordingCanvas();
        painter.Paint(canvas, sheet, viewport ?? Viewport, selection);
        return (canvas, painter);
    }

    private static GridSelection Range(string from, string to) =>
        GridSelection.Cell(Cells.Parse(from), null).ExtendTo(Cells.Parse(to), null);

    [Fact]
    public void The_border_goes_round_the_whole_selection()
    {
        var (canvas, painter) = Paint(Range("B2", "D5"));

        Assert.Equal(painter.CellBounds(Viewport, Cells.Parse("B2"), Cells.Parse("D5")), canvas.Outline(painter));
    }

    [Fact]
    public void Selected_cells_but_the_active_one_are_shaded()
    {
        var (canvas, painter) = Paint(Range("B2", "C3"));

        var shades = canvas.Calls.OfType<RecordingCanvas.Fill>().Where(f => f.Color == painter.Theme.SelectionFill).ToList();
        var active = painter.CellBounds(Viewport, Cells.Parse("B2"));
        var area = painter.CellBounds(Viewport, Cells.Parse("B2"), Cells.Parse("C3"));
        Assert.Equal(area.Width * area.Height - active.Width * active.Height, shades.Sum(s => s.Rect.Width * s.Rect.Height), 6);
        Assert.DoesNotContain(shades, s => s.Rect.X < active.Right && s.Rect.Right > active.X && s.Rect.Y < active.Bottom && s.Rect.Bottom > active.Y);
    }

    [Fact]
    public void A_single_cell_is_not_shaded()
    {
        var (canvas, painter) = Paint(GridSelection.Cell(Cells.Parse("B2"), null));

        Assert.DoesNotContain(canvas.Calls.OfType<RecordingCanvas.Fill>(), f => f.Color == painter.Theme.SelectionFill);
    }

    [Fact]
    public void Headers_of_whole_selected_columns_are_darker_and_coloured()
    {
        var (canvas, painter) = Paint(GridSelection.Columns(2, 3, 1, null));

        var theme = painter.Theme;
        Assert.Equal(theme.SelectedHeaderText, canvas.TextOf("B").Color);
        Assert.Equal(theme.SelectedHeaderText, canvas.TextOf("C").Color);
        Assert.Equal(theme.HeaderText, canvas.TextOf("D").Color);
        Assert.Equal(theme.HeaderText, canvas.TextOf("1").Color);
        var b = painter.CellBounds(Viewport, Cells.Parse("B1"));
        Assert.Contains(canvas.Calls.OfType<RecordingCanvas.Fill>(), f => f.Color == theme.SelectedHeaderBackground && f.Rect.X == b.X && f.Rect.Y == 0);
    }

    [Fact]
    public void Headers_of_a_selected_range_are_highlighted_and_underlined()
    {
        var (canvas, painter) = Paint(Range("B2", "C3"));

        var theme = painter.Theme;
        var fills = canvas.Calls.OfType<RecordingCanvas.Fill>().ToList();
        foreach (var column in new[] { "B1", "C1" })
            Assert.Contains(fills, f => f.Color == theme.ActiveHeaderBackground && f.Rect.X == painter.CellBounds(Viewport, Cells.Parse(column)).X && f.Rect.Y == 0);
        Assert.DoesNotContain(fills, f => f.Color == theme.ActiveHeaderBackground && f.Rect.X == painter.CellBounds(Viewport, Cells.Parse("D1")).X);
        Assert.Contains(canvas.Calls.OfType<RecordingCanvas.Line>(), l => l.Color == theme.ActiveCellBorder && l.Y1 == l.Y2 && l.Y1 == painter.ColumnHeaderHeight - 1);
    }

    [Fact]
    public void The_border_stops_at_the_frozen_line()
    {
        // B2:C10 with three rows frozen, scrolled so rows 4 to 6 are under the frozen part.
        var painter = new GridPainter { FrozenRows = 3 };
        var viewport = Viewport with { OffsetY = 3 * 20 };
        var (canvas, _) = Paint(Range("B2", "C10"), painter: painter, viewport: viewport);

        var frozenBottom = painter.ColumnHeaderHeight + 3 * 20;
        var lines = canvas.Calls.OfType<RecordingCanvas.Line>().Where(l => l.Color == painter.Theme.ActiveCellBorder && l.X1 == l.X2 && l.X1 > painter.RowHeaderWidth).ToList();

        // Each side is drawn twice, once in each pane, and never reaches across the other pane's part.
        Assert.Equal(4, lines.Count);
        Assert.All(lines, l => Assert.True(Math.Max(l.Y1, l.Y2) <= frozenBottom || Math.Min(l.Y1, l.Y2) >= frozenBottom - 1, $"{l.Y1}..{l.Y2}"));
    }

    [Fact]
    public void A_selection_without_the_sheet_still_covers_its_merged_areas()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet.Merge("B2:D3");

        var (canvas, painter) = Paint(GridSelection.Cell(Cells.Parse("B2"), null), sheet);

        Assert.Equal(painter.CellBounds(Viewport, Cells.Parse("B2"), Cells.Parse("D3")), canvas.Outline(painter));
    }

    [Fact]
    public void Hits_tell_cells_headers_and_the_corner_apart()
    {
        var painter = new GridPainter();
        var viewport = Viewport with { OffsetY = 100 * 20 };

        Assert.Equal(new GridHit(GridPart.Corner, Cells.Parse("A101")), painter.Hit(viewport, 5, 5));
        Assert.Equal(new GridHit(GridPart.ColumnHeader, Cells.Parse("B101")), painter.Hit(viewport, 40 + 64 + 10, 5));
        Assert.Equal(new GridHit(GridPart.RowHeader, Cells.Parse("A103")), painter.Hit(viewport, 5, 20 + 2 * 20 + 5));
        Assert.Equal(new GridHit(GridPart.Cell, Cells.Parse("B103")), painter.Hit(viewport, 40 + 64 + 10, 20 + 2 * 20 + 5));
        Assert.Null(painter.Hit(viewport, 700, 5));
    }

    [Fact]
    public void A_header_hit_with_frozen_panes_activates_the_first_frozen_cell()
    {
        var painter = new GridPainter { FrozenRows = 2 };

        var hit = painter.Hit(Viewport with { OffsetY = 1000 }, 40 + 10, 5);

        Assert.Equal(new GridHit(GridPart.ColumnHeader, Cells.Parse("A1")), hit);
    }

    [Fact]
    public void The_nearest_cell_past_an_edge_is_the_next_one_out_of_view()
    {
        var painter = new GridPainter();
        var viewport = Viewport with { OffsetX = 64, OffsetY = 20 };
        var visible = painter.VisibleArea(viewport);

        Assert.Equal(visible.LastColumn + 1, painter.CellNearest(viewport, 900, 100).Column);
        Assert.Equal(visible.LastRow + 1, painter.CellNearest(viewport, 100, 900).Row);
        Assert.Equal(new CellPosition(1, 1), painter.CellNearest(viewport, 0, 0));
        Assert.Equal(painter.HitTest(viewport, 200, 100), painter.CellNearest(viewport, 200, 100));
    }

    [Fact]
    public void Painting_the_active_cell_draws_what_its_selection_draws()
    {
        var sheet = new Workbook().AddSheet("S");
        var painter = new GridPainter();
        var byCell = new RecordingCanvas();
        var bySelection = new RecordingCanvas();

        painter.Paint(byCell, sheet, Viewport, Cells.Parse("C4"));
        painter.Paint(bySelection, sheet, Viewport, GridSelection.Cell(Cells.Parse("C4"), sheet));

        Assert.Equal(bySelection.Calls, byCell.Calls);
    }
}
