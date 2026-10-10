using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace SharpCell.Grid.Wpf.Tests;

public class SpreadsheetViewTests
{
    private const int Width = 480;
    private const int Height = 200;

    private static CellPosition Cell(string address)
    {
        var letters = address.TakeWhile(char.IsLetter).ToArray();
        var column = letters.Aggregate(0, (sum, letter) => sum * 26 + letter - 'A' + 1);
        return new CellPosition(int.Parse(address.AsSpan(letters.Length)), column);
    }

    private static Worksheet Budget()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Budget");
        sheet["A1"].Value = "Item";
        sheet["B1"].Value = "Price";
        sheet["A2"].Value = "Rent";
        sheet["B2"].Value = 1250;
        sheet["B2"].NumberFormat = "#,##0.00";
        sheet["D4"].Style = CellStyle.Default with { Fill = CellColor.FromRgb(0xFF0000) };
        workbook.Recalculate();
        return sheet;
    }

    private static SpreadsheetView Show(Worksheet? sheet)
    {
        var view = new SpreadsheetView { Worksheet = sheet };
        view.Measure(new Size(Width, Height));
        view.Arrange(new Rect(0, 0, Width, Height));
        view.UpdateLayout();
        return view;
    }

    private static Color[,] Render(SpreadsheetView view)
    {
        // A changed property only marks the view; layout paints it again before it is captured.
        view.UpdateLayout();
        var bitmap = new RenderTargetBitmap(Width, Height, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(view);
        var pixels = new byte[Width * Height * 4];
        bitmap.CopyPixels(pixels, Width * 4, 0);
        var colors = new Color[Width, Height];
        for (var y = 0; y < Height; y++)
        {
            for (var x = 0; x < Width; x++)
            {
                var i = (y * Width + x) * 4;
                colors[x, y] = Color.FromRgb(pixels[i + 2], pixels[i + 1], pixels[i]);
            }
        }

        return colors;
    }

    private static Point CenterOf(SpreadsheetView view, string address)
    {
        var bounds = view.Painter.CellBounds(view.Viewport, Cell(address));
        return new Point(bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2);
    }

    private static void Near(Color expected, Color actual, int tolerance = 8) =>
        Assert.True(Math.Abs(expected.R - actual.R) <= tolerance && Math.Abs(expected.G - actual.G) <= tolerance && Math.Abs(expected.B - actual.B) <= tolerance, $"expected {expected}, got {actual}");

    [Fact]
    public void Paints_headers_fills_and_text() => Sta.Run(() =>
    {
        var view = Show(Budget());

        var pixels = Render(view);

        Near(Color.FromRgb(0xF5, 0xF5, 0xF5), pixels[3, 3]);              // the corner
        var d4 = CenterOf(view, "D4");
        Near(Color.FromRgb(0xFF, 0x00, 0x00), pixels[(int)d4.X, (int)d4.Y]);
        var a2 = view.Painter.CellBounds(view.Viewport, Cell("A2"));
        var dark = 0;
        for (var x = (int)a2.X; x < (int)a2.Right; x++)
        {
            for (var y = (int)a2.Y; y < (int)a2.Bottom; y++)
                dark += pixels[x, y].R < 128 ? 1 : 0;
        }

        Assert.True(dark > 20, $"'Rent' should be drawn in A2: {dark} dark pixels");
    });

    [Fact]
    public void A_selected_range_is_shaded_but_for_the_active_cell() => Sta.Run(() =>
    {
        var view = Show(Budget());
        view.Selection = GridSelection.Cell(Cell("B3"), view.Worksheet).ExtendTo(Cell("C4"), view.Worksheet);

        var pixels = Render(view);

        var (b3, c4) = (CenterOf(view, "B3"), CenterOf(view, "C4"));
        Near(Colors.White, pixels[(int)b3.X, (int)b3.Y]);
        Near(Color.FromRgb(0xD9, 0xD9, 0xD9), pixels[(int)c4.X, (int)c4.Y], 12);
        var outline = view.Painter.CellBounds(view.Viewport, Cell("B3"), Cell("C4"));
        Near(Color.FromRgb(0x10, 0x7C, 0x41), pixels[(int)outline.X - 1, (int)b3.Y]);
        Near(Color.FromRgb(0x10, 0x7C, 0x41), pixels[(int)outline.X, (int)b3.Y]);
        Near(Colors.White, pixels[(int)outline.X + 1, (int)b3.Y]);
    });

    [Fact]
    public void Keys_move_and_extend_the_selection() => Sta.Run(() =>
    {
        var view = Show(Budget());
        view.ActiveCell = Cell("B2");

        view.PressKey(Key.Right, ModifierKeys.Shift);
        view.PressKey(Key.Down, ModifierKeys.Shift);
        Assert.Equal("B2:C3", view.Selection?.ToString());

        view.PressKey(Key.Enter, ModifierKeys.None);
        Assert.Equal(Cell("B3"), view.ActiveCell);

        view.PressKey(Key.Space, ModifierKeys.Control);
        Assert.Equal("B:C", view.Selection?.ToString());

        view.PressKey(Key.Left, ModifierKeys.None);
        Assert.Equal("A3", view.Selection?.ToString());
    });

    [Fact]
    public void Moving_right_far_enough_scrolls() => Sta.Run(() =>
    {
        var view = Show(Budget());

        for (var i = 0; i < 10; i++)
            view.PressKey(Key.Right, ModifierKeys.None);

        Assert.Equal(Cell("K1"), view.ActiveCell);
        Assert.True(view.Offset.X > 0);
    });

    [Fact]
    public void Clicking_and_dragging_selects_a_range() => Sta.Run(() =>
    {
        var view = Show(Budget());

        Assert.True(view.Press(CenterOf(view, "B2"), shift: false));
        view.DragTo(CenterOf(view, "D4"));

        Assert.Equal("B2:D4", view.Selection?.ToString());
        Assert.Equal(Cell("B2"), view.ActiveCell);
    });

    [Fact]
    public void A_column_header_selects_the_column() => Sta.Run(() =>
    {
        var view = Show(Budget());

        view.Press(new Point(CenterOf(view, "C1").X, 5), shift: false);

        Assert.Equal("C:C", view.Selection?.ToString());
    });

    [Fact]
    public void Select_all_and_copy_are_commands() => Sta.Run(() =>
    {
        var view = Show(Budget());
        view.Selection = GridSelection.Cell(Cell("A1"), view.Worksheet).ExtendTo(Cell("B2"), view.Worksheet);

        ApplicationCommands.Copy.Execute(null, view);
        Assert.Equal("Item\tPrice\r\nRent\t1,250.00\r\n", Clipboard.GetText());
        var html = Clipboard.GetData(DataFormats.Html) switch
        {
            string text => text,
            System.IO.MemoryStream stream => System.Text.Encoding.UTF8.GetString(stream.ToArray()),
            var other => other?.ToString() ?? "",
        };
        Assert.StartsWith("Version:0.9", html, StringComparison.Ordinal);
        Assert.Contains("<table", html, StringComparison.Ordinal);

        ApplicationCommands.SelectAll.Execute(null, view);
        Assert.Equal("1:1048576", view.Selection?.ToString());
    });

    [Fact]
    public void Setting_the_active_cell_selects_it_alone() => Sta.Run(() =>
    {
        var view = Show(Budget());
        view.Selection = GridSelection.Cell(Cell("B2"), view.Worksheet).Extend(2, 2, view.Worksheet);

        view.ActiveCell = Cell("A1");

        Assert.Equal("A1", view.Selection?.ToString());
    });
}
