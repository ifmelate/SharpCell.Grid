using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using SharpCell.Grid.Avalonia;

namespace SharpCell.Grid.Tests.Controls;

public class SheetEditingTests(HeadlessSession session)
{
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

    private static (Window Window, SpreadsheetView View) Show(Worksheet sheet)
    {
        var view = new SpreadsheetView { Worksheet = sheet };
        var window = new Window { Width = 480, Height = 240, Content = view };
        window.Show();
        view.Focus();
        return (window, view);
    }

    private static GridRect Bounds(SpreadsheetView view, string address) => view.Painter.CellBounds(view.Viewport, Cells.Parse(address));

    private static Point CenterOf(SpreadsheetView view, string address)
    {
        var bounds = Bounds(view, address);
        return new Point(bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2);
    }

    private static void Drag(Window window, Point from, Point to)
    {
        window.MouseDown(from, MouseButton.Left);
        window.MouseMove(to, RawInputModifiers.LeftMouseButton);
        window.MouseUp(to, MouseButton.Left);
    }

    private static void Press(Window window, Key key, RawInputModifiers modifiers = RawInputModifiers.None) =>
        window.KeyPress(key, modifiers, PhysicalKey.None, null);

    [Fact]
    public Task Dragging_a_column_header_edge_resizes_the_column_and_undo_gives_it_back() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        var a = Bounds(view, "A1");

        Drag(window, new Point(a.Right - 1, 10), new Point(a.Right + 35, 10));

        Assert.Equal(a.Width + 36, view.Painter.Layout.ColumnWidth(1));
        Assert.NotNull(sheet.ColumnWidth(1));
        view.Undo();
        Assert.Null(sheet.ColumnWidth(1));
        Assert.Equal(a.Width, view.Painter.Layout.ColumnWidth(1));
    });

    [Fact]
    public Task Dragging_a_row_edge_to_nothing_hides_the_row() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        var row = Bounds(view, "A2");

        Drag(window, new Point(5, row.Bottom - 1), new Point(5, row.Y - 30));

        Assert.True(sheet.IsRowHidden(2));
        Assert.Equal(0, view.Painter.Layout.RowHeight(2));
    });

    [Fact]
    public Task Double_clicking_a_column_edge_fits_the_column_to_its_text() => session.Run(() =>
    {
        var sheet = Sheet();
        sheet["C3"].Value = "a much longer piece of text";
        var (window, view) = Show(sheet);
        var c = Bounds(view, "C1");
        var edge = new Point(c.Right - 1, 10);

        window.MouseDown(edge, MouseButton.Left);
        window.MouseUp(edge, MouseButton.Left);
        window.MouseDown(edge, MouseButton.Left);
        window.MouseUp(edge, MouseButton.Left);

        Assert.Equal(view.Painter.FitColumnWidth(sheet, 3), view.Painter.Layout.ColumnWidth(3));
        Assert.True(view.Painter.Layout.ColumnWidth(3) > c.Width * 2);
    });

    [Fact]
    public Task Dragging_the_fill_handle_continues_the_series_and_selects_it() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.Selection = GridSelection.Cell(Cells.Parse("A1"), sheet).ExtendTo(Cells.Parse("A2"), sheet);
        var corner = Bounds(view, "A2");

        Drag(window, new Point(corner.Right - 1, corner.Bottom - 1), CenterOf(view, "A5"));

        Assert.Equal([3.0, 4.0, 5.0], new[] { "A3", "A4", "A5" }.Select(a => sheet[a].Value.AsNumber()));
        Assert.Equal("A1:A5", view.Selection?.ToString());
        Assert.Null(view.Painter.FillArea);
    });

    [Fact]
    public Task A_read_only_view_has_no_fill_handle_and_resizes_nothing() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.IsReadOnly = true;
        var a = Bounds(view, "A1");

        Drag(window, new Point(a.Right - 1, 10), new Point(a.Right + 35, 10));

        Assert.False(view.Painter.ShowsFillHandle);
        Assert.Null(sheet.ColumnWidth(1));
    });

    [Fact]
    public Task Ctrl_D_fills_the_top_row_down_with_formulas_following() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.Selection = GridSelection.Cell(Cells.Parse("B1"), sheet).ExtendTo(Cells.Parse("B2"), sheet);

        Press(window, Key.D, RawInputModifiers.Control);

        Assert.Equal("=A2*10", sheet["B2"].Formula);
        Assert.Equal(20.0, sheet["B2"].Value.AsNumber());
    });

    [Fact]
    public Task Ctrl_Plus_and_Ctrl_Minus_insert_and_delete_the_selected_rows() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("A1");
        Press(window, Key.Space, RawInputModifiers.Shift);

        Press(window, Key.OemPlus, RawInputModifiers.Control | RawInputModifiers.Shift);
        Assert.Equal("=A2*10", sheet["B2"].Formula);
        Assert.True(view.Selection!.IsEntireRows);

        Press(window, Key.OemMinus, RawInputModifiers.Meta);
        Assert.Equal("=A1*10", sheet["B1"].Formula);
    });

    [Fact]
    public Task Ctrl_Plus_and_Ctrl_Minus_with_cells_selected_ask_how_as_Excel_does() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("A1");

        Press(window, Key.OemPlus, RawInputModifiers.Control | RawInputModifiers.Shift);
        var dialog = Assert.IsType<SpreadsheetView.CellsDialog>(view.OpenDialog);
        Assert.Equal(0, dialog.Choice);                                  // one cell: shift cells right
        dialog.Choose(1);                                               // shift cells down
        Assert.Null(view.OpenDialog);
        Assert.Equal((CellValue.Empty, 1.0), (sheet["A1"].Value, sheet["A2"].Value.AsNumber()));
        Assert.Equal("=A2*10", sheet["B1"].Formula);                    // column B stays, its reference follows

        Press(window, Key.OemMinus, RawInputModifiers.Control);
        view.OpenDialog!.Choose(1);                                     // shift cells up
        Assert.Equal(1.0, sheet["A1"].Value.AsNumber());
        Assert.Equal("=A1*10", sheet["B1"].Formula);

        Press(window, Key.OemPlus, RawInputModifiers.Control | RawInputModifiers.Shift);
        view.OpenDialog!.Cancel();
        Assert.Equal(1.0, sheet["A1"].Value.AsNumber());
        view.Undo();
        Assert.Equal(CellValue.Empty, sheet["A1"].Value);              // undo takes back the delete
    });

    [Fact]
    public Task Ctrl_with_the_wheel_zooms_by_ten_percent() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);

        window.MouseWheel(CenterOf(view, "B2"), new Vector(0, 1), RawInputModifiers.Control);

        Assert.Equal(1.1, view.Zoom, 9);
        Assert.Equal(1.1, view.Painter.Zoom, 9);
        Assert.Equal(22, view.Painter.Layout.DefaultRowHeight);
        view.Zoom = 9;
        Assert.Equal(GridPainter.MaxZoom, view.Zoom);
    });

    [Fact]
    public Task Cells_copied_in_the_view_paste_with_their_formulas_until_Escape() => session.Run(async () =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("B1");

        await view.CopyAsync();
        Assert.Equal(new CellArea(1, 2, 1, 2), view.Painter.CopiedArea);
        view.Selection = GridSelection.Cell(Cells.Parse("B2"), sheet).ExtendTo(Cells.Parse("B3"), sheet);
        await view.PasteAsync();

        Assert.Equal("=A2*10", sheet["B2"].Formula);
        Assert.Equal("=A3*10", sheet["B3"].Formula);
        Assert.NotNull(view.Painter.CopiedArea);
        Press(window, Key.Escape);
        Assert.Null(view.Painter.CopiedArea);
    });

    [Fact]
    public Task Inserting_from_code_keeps_the_selection_and_takes_the_new_layout() => session.Run(() =>
    {
        var sheet = Sheet();
        sheet.SetRowHeight(2, 40);
        var (_, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("A2");

        view.InsertRows();

        Assert.Equal("A2", view.Selection?.ToString());
        Assert.Equal(Math.Round(40 * 96 / 72.0), view.Painter.Layout.RowHeight(3));
        view.HideColumns();
        Assert.True(sheet.IsColumnHidden(1));
        Assert.Equal(0, view.Painter.Layout.ColumnWidth(1));
    });

    [Fact]
    public Task A_touchpad_zooms_once_its_small_turns_make_a_whole_one() => session.Run(() =>
    {
        var (window, view) = Show(Sheet());

        for (var i = 0; i < 4; i++)
            window.MouseWheel(CenterOf(view, "B2"), new Vector(0, 0.2), RawInputModifiers.Meta);
        Assert.Equal(1, view.Zoom);
        window.MouseWheel(CenterOf(view, "B2"), new Vector(0, 0.2), RawInputModifiers.Meta);
        Assert.Equal(1.1, view.Zoom, 9);
    });

    [Fact]
    public Task A_right_click_on_a_column_header_selects_the_column_for_its_menu() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        var c = Bounds(view, "C1");

        window.MouseDown(new Point(c.X + c.Width / 2, 10), MouseButton.Right);
        window.MouseUp(new Point(c.X + c.Width / 2, 10), MouseButton.Right);

        Assert.True(view.Selection!.IsEntireColumns);
        Assert.Equal(3, view.Selection.Area.FirstColumn);
    });

    [Fact]
    public Task Pasting_over_the_whole_sheet_copies_once() => session.Run(async () =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("B1");
        await view.CopyAsync();

        Press(window, Key.A, RawInputModifiers.Control);
        Press(window, Key.A, RawInputModifiers.Control);   // the data around B1, then the sheet
        await view.PasteAsync();

        Assert.Equal("=A1*10", sheet["B1"].Formula);
        Assert.Null(sheet["B2"].Formula);
    });
}
