using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Platform;
using SharpCell.Grid.Avalonia;

namespace SharpCell.Grid.Tests.Controls;

public class SpreadsheetViewTests(HeadlessSession session)
{
    private static Worksheet Budget()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Budget");
        sheet["A1"].Value = "Item";
        sheet["B1"].Value = "Price";
        sheet["C1"].Value = "Share";
        sheet["A2"].Value = "Rent";
        sheet["B2"].Value = 1250;
        sheet["B2"].NumberFormat = "#,##0.00";
        sheet["A3"].Value = "Food";
        sheet["B3"].Value = 410.5;
        sheet["B3"].NumberFormat = "#,##0.00";
        sheet["A4"].Value = "Total";
        sheet["B4"].Formula = "=SUM(B2:B3)";
        sheet["B4"].NumberFormat = "#,##0.00";
        sheet["C2"].Formula = "=B2/$B$4";
        sheet["C3"].Formula = "=B3/$B$4";
        sheet["C2"].NumberFormat = "0.0%";
        sheet["C3"].NumberFormat = "0.0%";
        sheet["D2"].Value = true;
        sheet["D3"].Formula = "=1/0";
        sheet["E2"].Value = 45941;
        sheet["E2"].NumberFormat = "yyyy-mm-dd";
        workbook.Recalculate();
        return sheet;
    }

    private static (Window Window, SpreadsheetView View) Show(Worksheet? sheet, double width = 480, double height = 200)
    {
        var view = new SpreadsheetView { Worksheet = sheet };
        var window = new Window { Width = width, Height = height, Content = view };
        window.Show();
        view.Focus();
        return (window, view);
    }

    private static Point CenterOf(SpreadsheetView view, string address)
    {
        var cell = Cells.Parse(address);
        var bounds = view.Painter.CellBounds(view.Viewport, cell);
        return new Point(bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2);
    }

    [Fact]
    public Task Renders_a_sheet_as_approved() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        view.ActiveCell = Cells.Parse("B4");

        Snapshot.Match(window.CaptureRenderedFrame(), "budget");
    });

    private static Worksheet Styled()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Report");
        var header = CellStyle.Default with
        {
            Font = CellFont.Default with { Bold = true, Color = CellColor.FromRgb(0xFFFFFF) },
            Fill = CellColor.FromRgb(0x4472C4),
            HorizontalAlignment = CellHorizontalAlignment.Center,
            VerticalAlignment = CellVerticalAlignment.Center,
        };
        sheet["A1"].Value = "Quarterly report";
        sheet["A1"].Style = header with { Font = header.Font with { Size = 14 } };
        sheet.Merge("A1:D1");
        sheet.SetRowHeight(1, 24);
        sheet.SetColumnWidth(1, 16);
        string[] titles = ["Region", "Q1", "Q2", "Note"];
        for (var i = 0; i < titles.Length; i++)
        {
            sheet[2, i + 1].Value = titles[i];
            sheet[2, i + 1].Style = CellStyle.Default with
            {
                Font = CellFont.Default with { Bold = true },
                BottomBorder = new CellBorder(CellBorderStyle.Medium),
            };
        }

        sheet["A3"].Value = "North";
        sheet["B3"].Value = 1200;
        sheet["C3"].Value = 1350;
        sheet["A4"].Value = "South";
        sheet["A4"].Style = CellStyle.Default with { Font = CellFont.Default with { Italic = true, Color = CellColor.FromRgb(0xC00000) } };
        sheet["B4"].Value = 980;
        sheet["C4"].Value = 1010;
        sheet["B4"].Style = CellStyle.Default with { Fill = CellColor.FromRgb(0xFFF2CC) };
        sheet["D3"].Value = "Text that wraps in its cell";
        sheet["D3"].Style = CellStyle.Default with { WrapText = true, VerticalAlignment = CellVerticalAlignment.Top };
        sheet.SetRowHeight(3, 45);
        sheet["A5"].Value = "Total";
        sheet["A5"].Style = CellStyle.Default with { HorizontalAlignment = CellHorizontalAlignment.Right, Indent = 1 };
        sheet["B5"].Formula = "=SUM(B3:B4)";
        sheet["C5"].Formula = "=SUM(C3:C4)";
        foreach (var address in new[] { "B5", "C5" })
            sheet[address].Style = CellStyle.Default with
            {
                Font = CellFont.Default with { Underline = CellUnderline.Double },
                TopBorder = new CellBorder(CellBorderStyle.Thin),
                BottomBorder = new CellBorder(CellBorderStyle.Double),
            };
        sheet.SetColumnHidden(5, true);
        sheet["F2"].Value = "after a hidden column";
        workbook.Recalculate();
        return sheet;
    }

    [Fact]
    public Task Renders_styles_merges_and_sizes_as_approved() => session.Run(() =>
    {
        var (window, view) = Show(Styled(), width: 560, height: 220);
        view.ActiveCell = Cells.Parse("C4");

        Snapshot.Match(window.CaptureRenderedFrame(), "styles");
    });

    [Fact]
    public Task Clicking_inside_a_merge_selects_its_top_left_cell() => session.Run(() =>
    {
        var (window, view) = Show(Styled(), width: 560, height: 220);

        var point = CenterOf(view, "C1");
        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);

        Assert.Equal(Cells.Parse("A1"), view.ActiveCell);
    });

    [Fact]
    public Task Frozen_panes_stay_in_place_while_the_rest_scrolls() => session.Run(() =>
    {
        var sheet = Styled();
        sheet.FrozenRows = 2;
        sheet.FrozenColumns = 1;
        sheet.ShowGridlines = false;
        var (window, view) = Show(sheet, width: 560, height: 220);

        var point = CenterOf(view, "B4");
        for (var i = 0; i < 3; i++)
            window.MouseWheel(point, new Vector(0, -1));
        view.Offset = view.Offset.WithX(64);

        Assert.Equal(view.Painter.ColumnHeaderHeight, view.Painter.CellBounds(view.Viewport, Cells.Parse("A1")).Y);
        Snapshot.Match(window.CaptureRenderedFrame(), "frozen");
    });

    [Fact]
    public Task Renders_the_end_of_the_sheet_as_approved() => session.Run(() =>
    {
        var (window, view) = Show(null);
        view.Offset = new Vector(1e9, 1e9);

        Snapshot.Match(window.CaptureRenderedFrame(), "end-of-sheet");
    });

    [Fact]
    public Task Clicking_a_cell_makes_it_active() => session.Run(() =>
    {
        var (window, view) = Show(Budget());

        var point = CenterOf(view, "C3");
        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);

        Assert.Equal(Cells.Parse("C3"), view.ActiveCell);
    });

    [Fact]
    public Task Clicking_a_row_header_selects_the_row() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        view.ActiveCell = Cells.Parse("B2");

        window.MouseDown(new Point(5, 50), MouseButton.Left);
        window.MouseUp(new Point(5, 50), MouseButton.Left);

        Assert.Equal("2:2", view.Selection?.ToString());
        Assert.Equal(Cells.Parse("A2"), view.ActiveCell);
    });

    [Fact]
    public Task Arrow_keys_move_the_active_cell_and_scroll_it_into_view() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        view.ActiveCell = Cells.Parse("A1");

        for (var i = 0; i < 10; i++)
            window.KeyPress(Key.Right, RawInputModifiers.None, PhysicalKey.ArrowRight, null);
        window.KeyPress(Key.Down, RawInputModifiers.None, PhysicalKey.ArrowDown, null);

        Assert.Equal(Cells.Parse("K2"), view.ActiveCell);
        var bounds = view.Painter.CellBounds(view.Viewport, Cells.Parse("K2"));
        Assert.True(bounds.Right <= view.Bounds.Width, $"K2 ends at {bounds.Right}");
        Assert.True(view.Offset.X > 0);
    });

    [Fact]
    public Task Keys_stop_at_the_edge_of_the_sheet() => session.Run(() =>
    {
        var (window, view) = Show(null);
        view.ActiveCell = Cells.Parse("A1");

        window.KeyPress(Key.Left, RawInputModifiers.None, PhysicalKey.ArrowLeft, null);
        window.KeyPress(Key.Up, RawInputModifiers.None, PhysicalKey.ArrowUp, null);

        Assert.Equal(Cells.Parse("A1"), view.ActiveCell);
    });

    [Fact]
    public Task Enter_moves_down_and_Tab_right_as_in_Excel() => session.Run(() =>
    {
        var (window, view) = Show(null);
        view.ActiveCell = Cells.Parse("B2");

        window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, null);
        window.KeyPress(Key.Tab, RawInputModifiers.None, PhysicalKey.Tab, null);

        Assert.Equal(Cells.Parse("C3"), view.ActiveCell);
    });

    [Fact]
    public Task The_wheel_scrolls_three_rows_and_never_above_the_top() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        var point = CenterOf(view, "B2");

        window.MouseWheel(point, new Vector(0, -1));
        Assert.Equal(3 * SheetLayout.ExcelRowHeight, view.Offset.Y);

        window.MouseWheel(point, new Vector(0, 2));
        Assert.Equal(0, view.Offset.Y);
    });

    [Fact]
    public Task Dragging_selects_a_range_from_where_it_started() => session.Run(() =>
    {
        var (window, view) = Show(Budget());

        window.MouseDown(CenterOf(view, "B2"), MouseButton.Left);
        window.MouseMove(CenterOf(view, "D4"), RawInputModifiers.LeftMouseButton);
        window.MouseUp(CenterOf(view, "D4"), MouseButton.Left);

        Assert.Equal("B2:D4", view.Selection?.ToString());
        Assert.Equal(Cells.Parse("B2"), view.ActiveCell);
    });

    [Fact]
    public Task Dragging_past_the_edge_scrolls_after_the_mouse() => session.Run(() =>
    {
        var (window, view) = Show(Budget());

        window.MouseDown(CenterOf(view, "B2"), MouseButton.Left);
        for (var i = 0; i < 5; i++)
            window.MouseMove(new Point(100, 400 + i), RawInputModifiers.LeftMouseButton);
        window.MouseUp(new Point(100, 405), MouseButton.Left);

        // Nine rows fit in the view; each move past its bottom edge takes one more.
        Assert.True(view.Offset.Y > 0);
        Assert.True(view.Selection!.Area.LastRow >= 14, view.Selection.ToString());
        Assert.Equal(Cells.Parse("B2"), view.ActiveCell);
    });

    [Fact]
    public Task Moving_the_mouse_without_a_button_selects_nothing() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        window.MouseDown(CenterOf(view, "B2"), MouseButton.Left);
        window.MouseUp(CenterOf(view, "B2"), MouseButton.Left);

        window.MouseMove(CenterOf(view, "D4"));

        Assert.Equal("B2", view.Selection?.ToString());
    });

    [Fact]
    public Task Shift_click_extends_from_the_active_cell() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        view.ActiveCell = Cells.Parse("B2");

        window.MouseDown(CenterOf(view, "C4"), MouseButton.Left, RawInputModifiers.Shift);
        window.MouseUp(CenterOf(view, "C4"), MouseButton.Left, RawInputModifiers.Shift);

        Assert.Equal("B2:C4", view.Selection?.ToString());
    });

    [Fact]
    public Task Dragging_over_column_headers_selects_whole_columns() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        var (b, d) = (CenterOf(view, "B1"), CenterOf(view, "D1"));

        window.MouseDown(new Point(b.X, 5), MouseButton.Left);
        window.MouseMove(new Point(d.X, 5), RawInputModifiers.LeftMouseButton);
        window.MouseUp(new Point(d.X, 5), MouseButton.Left);

        Assert.Equal("B:D", view.Selection?.ToString());
        Assert.Equal(Cells.Parse("B1"), view.ActiveCell);
    });

    [Fact]
    public Task Shift_with_the_arrows_extends_and_an_arrow_alone_collapses() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        view.ActiveCell = Cells.Parse("B2");

        window.KeyPress(Key.Right, RawInputModifiers.Shift, PhysicalKey.ArrowRight, null);
        window.KeyPress(Key.Down, RawInputModifiers.Shift, PhysicalKey.ArrowDown, null);
        Assert.Equal("B2:C3", view.Selection?.ToString());

        window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, null);
        Assert.Equal(Cells.Parse("B3"), view.ActiveCell);
        Assert.Equal("B2:C3", view.Selection?.ToString());

        window.KeyPress(Key.Left, RawInputModifiers.None, PhysicalKey.ArrowLeft, null);
        Assert.Equal("A3", view.Selection?.ToString());
    });

    [Fact]
    public Task Ctrl_space_shift_space_and_select_all() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        view.ActiveCell = Cells.Parse("B2");

        window.KeyPress(Key.Space, RawInputModifiers.Control, PhysicalKey.Space, " ");
        Assert.Equal("B:B", view.Selection?.ToString());

        window.KeyPress(Key.Space, RawInputModifiers.Shift, PhysicalKey.Space, " ");
        Assert.Equal("1:1048576", view.Selection?.ToString());

        view.ActiveCell = Cells.Parse("C3");
        window.KeyPress(Key.A, RawInputModifiers.Control, PhysicalKey.A, "a");
        Assert.Equal("1:1048576", view.Selection?.ToString());
        Assert.Equal(Cells.Parse("C3"), view.ActiveCell);
    });

    [Fact]
    public Task Setting_the_active_cell_selects_it_alone() => session.Run(() =>
    {
        var (_, view) = Show(Budget());
        view.Selection = GridSelection.Cell(Cells.Parse("B2"), view.Worksheet).Extend(2, 2, view.Worksheet);

        view.ActiveCell = Cells.Parse("A1");

        Assert.Equal("A1", view.Selection?.ToString());
    });

    [Fact]
    public Task Copying_puts_the_selected_cells_on_the_clipboard() => session.Run(async () =>
    {
        var (window, view) = Show(Budget());
        view.Selection = GridSelection.Cell(Cells.Parse("A1"), view.Worksheet).ExtendTo(Cells.Parse("B2"), view.Worksheet);

        await view.CopyAsync();

        var clipboard = window.Clipboard!;
        Assert.Equal("Item\tPrice\r\nRent\t1,250.00\r\n", await clipboard.TryGetTextAsync());
    });

    [Theory]
    [InlineData(RawInputModifiers.Control)]
    [InlineData(RawInputModifiers.Meta)]   // Cmd on a Mac, also in a browser, whose gestures use Ctrl
    public Task Ctrl_c_and_cmd_c_copy(RawInputModifiers modifier) => session.Run(async () =>
    {
        var (window, view) = Show(Budget());
        view.Selection = GridSelection.Cell(Cells.Parse("A1"), view.Worksheet).ExtendTo(Cells.Parse("B1"), view.Worksheet);
        await window.Clipboard!.SetTextAsync("before");

        window.KeyPress(Key.C, modifier, PhysicalKey.C, "c");
        await Task.Delay(50);

        Assert.Equal("Item\tPrice\r\n", await window.Clipboard!.TryGetTextAsync());
    });

    [Theory]
    [InlineData(RawInputModifiers.Control)]
    [InlineData(RawInputModifiers.Meta)]
    public Task Ctrl_a_and_cmd_a_select_the_whole_sheet(RawInputModifiers modifier) => session.Run(() =>
    {
        var (window, view) = Show(Budget());

        window.KeyPress(Key.A, modifier, PhysicalKey.A, "a");

        Assert.Equal("1:1048576", view.Selection?.ToString());
    });

    [Fact]
    public Task Renders_a_selection_as_approved() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        view.Selection = GridSelection.Cell(Cells.Parse("B2"), view.Worksheet).ExtendTo(Cells.Parse("C4"), view.Worksheet);

        Snapshot.Match(window.CaptureRenderedFrame(), "selection");
    });

    [Fact]
    public Task Renders_whole_columns_as_approved() => session.Run(() =>
    {
        var (window, view) = Show(Budget());
        view.Selection = GridSelection.Columns(2, 3, 1, view.Worksheet);

        Snapshot.Match(window.CaptureRenderedFrame(), "columns");
    });
}
