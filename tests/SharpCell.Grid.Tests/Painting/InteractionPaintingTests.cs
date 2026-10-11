using System.Globalization;

namespace SharpCell.Grid.Tests.Painting;

public class InteractionPaintingTests
{
    private static readonly GridViewport Viewport = new(0, 0, 600, 300);

    // Seven pixels a character at Calibri 11's 14.67 pixels, growing with the font.
    private static double Measure(string text, GridFont font) => text.Length * 7 * font.Size / (11 * 96 / 72.0);

    [Fact]
    public void Edges_between_column_headers_are_found_three_pixels_either_side()
    {
        var painter = new GridPainter();

        // Column A ends at 40 + 64 = 104.
        Assert.Equal(1, painter.ColumnEdgeAt(Viewport, 102, 10));
        Assert.Equal(1, painter.ColumnEdgeAt(Viewport, 106, 10));
        Assert.Null(painter.ColumnEdgeAt(Viewport, 110, 10));
        Assert.Null(painter.ColumnEdgeAt(Viewport, 103, 30));
        Assert.Equal(1, painter.RowEdgeAt(Viewport, 10, 39));
        Assert.Equal(1, painter.RowEdgeAt(Viewport, 10, 41));
        Assert.Null(painter.RowEdgeAt(Viewport, 50, 39));
    }

    [Fact]
    public void Right_of_hidden_columns_the_edge_is_the_visible_column_before_them()
    {
        var painter = new GridPainter();
        painter.Layout.SetColumnWidth(2, 0);

        Assert.Equal(1, painter.ColumnEdgeAt(Viewport, 106, 10));
        Assert.Equal(3, painter.ColumnEdgeAt(Viewport, 167, 10));
    }

    [Fact]
    public void The_fill_handle_sits_on_the_selection_corner_when_shown()
    {
        var painter = new GridPainter { ShowsFillHandle = true };
        var selection = GridSelection.Cell(Cells.Parse("B2"), null);
        var canvas = new RecordingCanvas();

        painter.Paint(canvas, null, Viewport, selection);

        // B2 is 104 to 168 by 40 to 60; its border's last pixels are 167 and 59.
        Assert.Contains(new RecordingCanvas.Fill(new GridRect(165, 57, 5, 5), painter.Theme.ActiveCellBorder), canvas.Calls);
        Assert.True(painter.IsOnFillHandle(Viewport, 168, 60, selection, null));
        Assert.False(painter.IsOnFillHandle(Viewport, 160, 59, selection, null));

        painter.ShowsFillHandle = false;
        Assert.False(painter.IsOnFillHandle(Viewport, 168, 60, selection, null));
    }

    [Fact]
    public void Copied_and_filled_cells_get_dashed_borders()
    {
        var painter = new GridPainter { CopiedArea = new CellArea(2, 2, 3, 3), FillArea = new CellArea(5, 1, 6, 1) };
        var canvas = new RecordingCanvas();

        painter.Paint(canvas, null, Viewport, (GridSelection?)null);

        var dashed = canvas.Calls.OfType<RecordingCanvas.Line>().Where(l => l.Dash == GridDash.Dash).ToList();
        Assert.Equal(4, dashed.Count(l => l.Color == painter.Theme.ActiveCellBorder));
        Assert.Equal(4, dashed.Count(l => l.Color == painter.Theme.FillOutline));
    }

    [Fact]
    public void Zoom_scales_fonts_headers_and_digits_but_not_lines()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = "x";
        var painter = new GridPainter { Zoom = 2 };
        var canvas = new RecordingCanvas();

        painter.Paint(canvas, sheet, Viewport, (GridSelection?)null);

        Assert.Equal(11 * 96 / 72.0 * 2, canvas.TextOf("x").Font.Size, 6);
        Assert.Equal(new GridRect(0, 0, 80, 40), canvas.TextOf("A").Bounds with { X = 0, Width = 80 });
        Assert.Equal(14, painter.MaxDigitWidth(sheet));
        Assert.All(canvas.Calls.OfType<RecordingCanvas.Line>().Where(l => l.Color == painter.Theme.Gridline), l => Assert.Equal(1, l.Thickness));
        Assert.Throws<ArgumentOutOfRangeException>(() => painter.Zoom = 5);
    }

    [Fact]
    public void A_column_fits_its_widest_text_and_a_row_its_tallest()
    {
        var workbook = new Workbook { Culture = CultureInfo.GetCultureInfo("en-US") };
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = "abcdefghij";
        sheet["A2"].Value = "abc";
        sheet["A3"].Value = "a much longer text in a merged cell";
        sheet.Merge("A3:B3");
        sheet["C1"].Value = "one\ntwo";
        sheet["D1"].Value = "aaa bbb ccc";
        sheet["D1"].Style = CellStyle.Default with { WrapText = true };
        sheet["A5"].Value = "big";
        sheet["A5"].Style = CellStyle.Default with { Font = CellFont.Default with { Size = 22 } };
        var painter = new GridPainter { MeasureText = Measure };

        Assert.Equal(75, painter.FitColumnWidth(sheet, 1));
        Assert.Null(painter.FitColumnWidth(sheet, 5));
        Assert.Equal(40, painter.FitRowHeight(sheet, 1));
        Assert.Equal(39, painter.FitRowHeight(sheet, 5));
        Assert.Null(painter.FitRowHeight(sheet, 4));
    }

    [Fact]
    public void Fitting_through_the_editor_is_one_step_in_Excel_units()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = "abcdefghij";
        sheet["B1"].Value = "abc";
        var editor = new GridEditor(workbook);
        var painter = new GridPainter { MeasureText = Measure };

        editor.FitColumnWidth(sheet, 1, 3, painter);

        Assert.Equal(75, SheetLayout.FromWorksheet(sheet, 7).ColumnWidth(1));
        Assert.Equal(26, SheetLayout.FromWorksheet(sheet, 7).ColumnWidth(2));
        Assert.Null(sheet.ColumnWidth(3));
        editor.Undo();
        Assert.Null(sheet.ColumnWidth(1));
    }
}
