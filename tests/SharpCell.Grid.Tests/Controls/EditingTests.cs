using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.LogicalTree;
using SharpCell.Grid.Avalonia;

namespace SharpCell.Grid.Tests.Controls;

public class EditingTests(HeadlessSession session)
{
    private static Worksheet Sheet()
    {
        var workbook = new Workbook { Culture = System.Globalization.CultureInfo.GetCultureInfo("en-US") };
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 2;
        sheet["B1"].Formula = "=A1*10";
        workbook.Recalculate();
        return sheet;
    }

    private static (Window Window, SpreadsheetView View) Show(Worksheet sheet)
    {
        var view = new SpreadsheetView { Worksheet = sheet };
        var window = new Window { Width = 480, Height = 200, Content = view };
        window.Show();
        view.Focus();
        return (window, view);
    }

    private static Point CenterOf(SpreadsheetView view, string address)
    {
        var bounds = view.Painter.CellBounds(view.Viewport, Cells.Parse(address));
        return new Point(bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2);
    }

    private static void Press(Window window, Key key, RawInputModifiers modifiers = RawInputModifiers.None) =>
        window.KeyPress(key, modifiers, PhysicalKey.None, null);

    [Fact]
    public Task Typing_starts_editing_and_Enter_enters_it_and_moves_down() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);

        window.KeyTextInput("5");
        Assert.True(view.IsEditing);
        Assert.Equal("5", view.EditText);
        Press(window, Key.Enter);

        Assert.False(view.IsEditing);
        Assert.Equal(5.0, sheet["A1"].Value.AsNumber());
        Assert.Equal(50.0, sheet["B1"].Value.AsNumber());
        Assert.Equal(Cells.Parse("A2"), view.ActiveCell);
    });

    [Fact]
    public Task Escape_leaves_the_cell_as_it_was() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);

        window.KeyTextInput("9");
        Press(window, Key.Escape);

        Assert.False(view.IsEditing);
        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
        Assert.Equal(Cells.Parse("A1"), view.ActiveCell);
    });

    [Fact]
    public Task F2_edits_the_formula_and_the_arrows_stay_in_the_text() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("B1");

        Press(window, Key.F2);
        Assert.Equal("=A1*10", view.EditText);
        Press(window, Key.Left);
        Assert.True(view.IsEditing);
        window.KeyTextInput("0");
        Press(window, Key.Tab);

        Assert.Equal("=A1*100", sheet["B1"].Formula);
        Assert.Equal(200.0, sheet["B1"].Value.AsNumber());
        Assert.Equal(Cells.Parse("C1"), view.ActiveCell);
    });

    [Fact]
    public Task An_arrow_after_typing_enters_and_moves() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("C3");

        window.KeyTextInput("hi");
        Press(window, Key.Right);

        Assert.Equal(CellValue.Text("hi"), sheet["C3"].Value);
        Assert.Equal(Cells.Parse("D3"), view.ActiveCell);
    });

    [Fact]
    public Task Ctrl_Enter_fills_the_selection() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.Selection = GridSelection.Cell(Cells.Parse("C1"), sheet).ExtendTo(Cells.Parse("D2"), sheet);

        window.KeyTextInput("7");
        Press(window, Key.Enter, RawInputModifiers.Control);

        foreach (var address in new[] { "C1", "D1", "C2", "D2" })
            Assert.Equal(7.0, sheet[address].Value.AsNumber());
        Assert.Equal("C1:D2", view.Selection?.ToString());
    });

    [Fact]
    public Task A_formula_that_does_not_parse_keeps_the_editor_open() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);

        window.KeyTextInput("=1+");
        Press(window, Key.Enter);

        Assert.True(view.IsEditing);
        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
        Press(window, Key.Escape);
        Assert.False(view.IsEditing);
    });

    [Fact]
    public Task Delete_clears_and_undo_and_redo_go_back_and_forth() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.Selection = GridSelection.Cell(Cells.Parse("A1"), sheet).ExtendTo(Cells.Parse("B1"), sheet);

        Press(window, Key.Delete);
        Assert.Equal(CellValue.Empty, sheet["A1"].Value);
        Assert.Null(sheet["B1"].Formula);

        view.ActiveCell = Cells.Parse("E5");
        Press(window, Key.Z, RawInputModifiers.Meta);
        Assert.Equal(20.0, sheet["B1"].Value.AsNumber());
        Assert.Equal("A1:B1", view.Selection?.ToString());

        Press(window, Key.Y, RawInputModifiers.Control);
        Assert.Equal(CellValue.Empty, sheet["A1"].Value);
    });

    [Fact]
    public Task Backspace_starts_with_an_empty_cell() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);

        Press(window, Key.Back);
        Assert.Equal("", view.EditText);
        Press(window, Key.Enter);

        Assert.Equal(CellValue.Empty, sheet["A1"].Value);
    });

    [Fact]
    public Task Pasting_puts_clipboard_text_in_cell_by_cell() => session.Run(async () =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("C2");
        await window.Clipboard!.SetTextAsync("1\t2\r\n3\t=C2+C3\r\n");

        await view.PasteAsync();

        Assert.Equal(1.0, sheet["C2"].Value.AsNumber());
        Assert.Equal(4.0, sheet["D3"].Value.AsNumber());
        Assert.Equal("C2:D3", view.Selection?.ToString());
    });

    [Fact]
    public Task A_read_only_view_changes_nothing() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.IsReadOnly = true;

        window.KeyTextInput("5");
        Press(window, Key.Delete);

        Assert.False(view.IsEditing);
        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
    });

    [Fact]
    public Task A_double_click_edits_and_a_click_elsewhere_enters() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);

        window.MouseDown(CenterOf(view, "A1"), MouseButton.Left);
        window.MouseUp(CenterOf(view, "A1"), MouseButton.Left);
        window.MouseDown(CenterOf(view, "A1"), MouseButton.Left);
        window.MouseUp(CenterOf(view, "A1"), MouseButton.Left);
        Assert.True(view.IsEditing);
        Assert.Equal("2", view.EditText);
        view.EditText = "3";

        window.MouseDown(CenterOf(view, "D4"), MouseButton.Left);
        window.MouseUp(CenterOf(view, "D4"), MouseButton.Left);

        Assert.False(view.IsEditing);
        Assert.Equal(3.0, sheet["A1"].Value.AsNumber());
        Assert.Equal(Cells.Parse("D4"), view.ActiveCell);
    });

    [Fact]
    public Task Changes_are_announced() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        var edits = new List<GridEdit>();
        view.Edited += (_, e) => edits.Add(e);

        window.KeyTextInput("4");
        Press(window, Key.Enter);

        var edit = Assert.Single(edits);
        Assert.Contains(edit.Recalculation.ChangedCells, c => c.Address == "B1");
    });

    [Fact]
    public Task Renders_a_cell_being_edited_as_approved() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view) = Show(sheet);
        view.ActiveCell = Cells.Parse("B2");

        window.KeyTextInput("A longer text than the cell");

        Snapshot.Match(window.CaptureRenderedFrame(), "editing");
    });

    private static (Window Window, SpreadsheetView View, FormulaBar Bar) ShowWithBar(Worksheet sheet)
    {
        var view = new SpreadsheetView { Worksheet = sheet };
        var bar = new FormulaBar { View = view };
        DockPanel.SetDock(bar, Dock.Top);
        var window = new Window { Width = 480, Height = 240, Content = new DockPanel { Children = { bar, view } } };
        window.Show();
        view.Focus();
        return (window, view, bar);
    }

    private static TextBox[] Boxes(FormulaBar bar) => bar.GetLogicalDescendants().OfType<TextBox>().ToArray();

    [Fact]
    public Task The_formula_bar_shows_the_active_cell_as_it_is_edited() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view, bar) = ShowWithBar(sheet);
        var (name, text) = (Boxes(bar)[0], Boxes(bar)[1]);

        view.ActiveCell = Cells.Parse("B1");
        Assert.Equal("B1", name.Text);
        Assert.Equal("=A1*10", text.Text);

        window.KeyTextInput("12");
        Assert.Equal("12", text.Text);
    });

    [Fact]
    public Task Typing_in_the_formula_bar_edits_the_cell() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view, bar) = ShowWithBar(sheet);
        var text = Boxes(bar)[1];
        view.ActiveCell = Cells.Parse("C1");

        text.Focus();
        Assert.True(view.IsEditing);
        text.Text = "=A1+1";
        Assert.Equal("=A1+1", view.EditText);
        window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, null);

        Assert.Equal("=A1+1", sheet["C1"].Formula);
        Assert.Equal(3.0, sheet["C1"].Value.AsNumber());
        Assert.Equal(Cells.Parse("C2"), view.ActiveCell);
    });

    [Fact]
    public Task An_address_in_the_name_box_selects_it() => session.Run(() =>
    {
        var sheet = Sheet();
        var (window, view, bar) = ShowWithBar(sheet);
        var name = Boxes(bar)[0];

        name.Focus();
        name.Text = "C40:D41";
        window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.Enter, null);

        Assert.Equal("C40:D41", view.Selection?.ToString());
        Assert.True(view.Offset.Y > 0);
    });

    [Fact]
    public Task Ctrl_arrows_home_end_and_pages_move_as_in_Excel() => session.Run(() =>
    {
        var sheet = Sheet();
        sheet["A5"].Value = 1;
        var (window, view) = Show(sheet);

        Press(window, Key.Down, RawInputModifiers.Control);
        Assert.Equal(Cells.Parse("A5"), view.ActiveCell);
        Press(window, Key.Right, RawInputModifiers.Meta | RawInputModifiers.Shift);
        Assert.Equal("5:5", view.Selection?.ToString());   // A5:XFD5, every column of the row

        Press(window, Key.End, RawInputModifiers.Control);
        Assert.Equal(Cells.Parse("B5"), view.ActiveCell);
        Press(window, Key.Home);
        Assert.Equal(Cells.Parse("A5"), view.ActiveCell);

        Press(window, Key.PageDown);
        Assert.True(view.ActiveCell!.Value.Row > 10);
        Assert.True(view.Offset.Y > 0);
        Press(window, Key.Home, RawInputModifiers.Control);
        Assert.Equal(Cells.Parse("A1"), view.ActiveCell);
        Assert.Equal(0, view.Offset.Y);
    });
}
