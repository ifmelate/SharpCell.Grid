using System.Windows;
using System.Windows.Input;

namespace SharpCell.Grid.Wpf.Tests;

public class SheetEditingTests
{
    private static CellPosition Cell(string address)
    {
        var letters = address.TakeWhile(char.IsLetter).ToArray();
        var column = letters.Aggregate(0, (sum, letter) => sum * 26 + letter - 'A' + 1);
        return new CellPosition(int.Parse(address.AsSpan(letters.Length)), column);
    }

    private static Worksheet Sheet()
    {
        var workbook = new Workbook { Culture = System.Globalization.CultureInfo.GetCultureInfo("en-US") };
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 1;
        sheet["A2"].Value = 2;
        sheet["B1"].Formula = "=A1*10";
        workbook.Recalculate();
        return sheet;
    }

    private static SpreadsheetView Show(Worksheet sheet)
    {
        var view = new SpreadsheetView { Worksheet = sheet };
        view.Measure(new Size(480, 240));
        view.Arrange(new Rect(0, 0, 480, 240));
        view.UpdateLayout();
        return view;
    }

    private static GridRect Bounds(SpreadsheetView view, string address) => view.Painter.CellBounds(view.Viewport, Cell(address));

    private static Point CenterOf(SpreadsheetView view, string address)
    {
        var bounds = Bounds(view, address);
        return new Point(bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2);
    }

    [Fact]
    public void Dragging_a_column_header_edge_resizes_the_column_and_undo_gives_it_back() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        var a = Bounds(view, "A1");

        Assert.True(view.Press(new Point(a.Right - 1, 10), shift: false));
        view.DragTo(new Point(a.Right + 35, 10));
        view.Release();

        Assert.Equal(a.Width + 36, view.Painter.Layout.ColumnWidth(1));
        Assert.NotNull(sheet.ColumnWidth(1));
        view.Undo();
        Assert.Null(sheet.ColumnWidth(1));
        Assert.Equal(a.Width, view.Painter.Layout.ColumnWidth(1));
    });

    [Fact]
    public void Double_clicking_a_column_edge_fits_the_column_to_its_text() => Sta.Run(() =>
    {
        var sheet = Sheet();
        sheet["C3"].Value = "a much longer piece of text";
        var view = Show(sheet);
        var c = Bounds(view, "C1");

        view.Press(new Point(c.Right - 1, 10), shift: false, clicks: 2);

        Assert.Equal(view.Painter.FitColumnWidth(sheet, 3), view.Painter.Layout.ColumnWidth(3));
        Assert.True(view.Painter.Layout.ColumnWidth(3) > c.Width * 2);
    });

    [Fact]
    public void Dragging_the_fill_handle_continues_the_series_and_selects_it() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.Selection = GridSelection.Cell(Cell("A1"), sheet).ExtendTo(Cell("A2"), sheet);
        var corner = Bounds(view, "A2");

        Assert.True(view.Press(new Point(corner.Right - 1, corner.Bottom - 1), shift: false));
        view.DragTo(CenterOf(view, "A5"));
        Assert.Equal(new CellArea(1, 1, 5, 1), view.Painter.FillArea);
        view.Release();

        Assert.Equal([3.0, 4.0, 5.0], new[] { "A3", "A4", "A5" }.Select(a => sheet[a].Value.AsNumber()));
        Assert.Equal("A1:A5", view.Selection?.ToString());
        Assert.Null(view.Painter.FillArea);
    });

    [Fact]
    public void Ctrl_D_and_Ctrl_Plus_fill_down_and_insert_rows() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.Selection = GridSelection.Cell(Cell("B1"), sheet).ExtendTo(Cell("B2"), sheet);

        Assert.True(view.PressKey(Key.D, ModifierKeys.Control));
        Assert.Equal("=A2*10", sheet["B2"].Formula);

        view.Selection = GridSelection.Rows(1, 1, 1, sheet);
        Assert.True(view.PressKey(Key.OemPlus, ModifierKeys.Control | ModifierKeys.Shift));
        Assert.Equal("=A2*10", sheet["B2"].Formula);
        Assert.True(view.PressKey(Key.OemMinus, ModifierKeys.Control));
        Assert.Equal("=A1*10", sheet["B1"].Formula);
    });

    [Fact]
    public void Ctrl_Plus_with_cells_selected_asks_how_as_Excel_does() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.ActiveCell = Cell("A1");

        Assert.True(view.PressKey(Key.OemPlus, ModifierKeys.Control | ModifierKeys.Shift));
        var dialog = Assert.IsType<SpreadsheetView.CellsDialog>(view.OpenDialog);
        dialog.Choose(1);

        Assert.Null(view.OpenDialog);
        Assert.Equal(CellValue.Empty, sheet["A1"].Value);
        Assert.Equal("=A2*10", sheet["B1"].Formula);
        view.Undo();
        Assert.Equal("=A1*10", sheet["B1"].Formula);
    });

    [Fact]
    public void Formatting_shortcuts_change_the_selected_cells_as_in_Excel() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.Selection = GridSelection.Cell(Cell("A1"), sheet).ExtendTo(Cell("A2"), sheet);

        Assert.True(view.PressKey(Key.B, ModifierKeys.Control));
        Assert.True(sheet["A2"].Style.Font.Bold);
        Assert.True(view.PressKey(Key.D5, ModifierKeys.Control | ModifierKeys.Shift));
        Assert.Equal("100%", sheet["A1"].Text);
        view.SetBorders(BorderEdges.Outside, new CellBorder(CellBorderStyle.Thin));
        Assert.NotNull(sheet["A2"].Style.BottomBorder);
        view.Undo();
        Assert.Null(sheet["A2"].Style.BottomBorder);
    });

    [Fact]
    public void Ctrl_with_the_wheel_zooms_by_ten_percent() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);

        view.Wheel(120, control: true);

        Assert.Equal(1.1, view.Zoom, 9);
        Assert.Equal(22, view.Painter.Layout.DefaultRowHeight);
        view.Zoom = 9;
        Assert.Equal(GridPainter.MaxZoom, view.Zoom);
    });

    [Fact]
    public void Cells_copied_in_the_view_paste_with_their_formulas_until_Escape() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.ActiveCell = Cell("B1");

        view.Copy();
        Assert.Equal(new CellArea(1, 2, 1, 2), view.Painter.CopiedArea);
        view.Selection = GridSelection.Cell(Cell("B2"), sheet).ExtendTo(Cell("B3"), sheet);
        view.Paste();

        Assert.Equal("=A3*10", sheet["B3"].Formula);
        Assert.True(view.PressKey(Key.Escape, ModifierKeys.None));
        Assert.Null(view.Painter.CopiedArea);
    });

    [Fact]
    public void A_read_only_view_has_no_fill_handle_and_resizes_nothing() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.IsReadOnly = true;
        var a = Bounds(view, "A1");

        view.Press(new Point(a.Right - 1, 10), shift: false);
        view.DragTo(new Point(a.Right + 35, 10));
        view.Release();

        Assert.False(view.Painter.ShowsFillHandle);
        Assert.Null(sheet.ColumnWidth(1));
    });
}
