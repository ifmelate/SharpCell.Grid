using System.Windows;
using System.Windows.Input;

namespace SharpCell.Grid.Wpf.Tests;

public class EditingTests
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
        sheet["A1"].Value = 2;
        sheet["B1"].Formula = "=A1*10";
        workbook.Recalculate();
        return sheet;
    }

    private static SpreadsheetView Show(Worksheet sheet)
    {
        var view = new SpreadsheetView { Worksheet = sheet };
        view.Measure(new Size(480, 200));
        view.Arrange(new Rect(0, 0, 480, 200));
        view.UpdateLayout();
        return view;
    }

    [Fact]
    public void Typing_and_Enter_enter_the_text_and_move_down() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);

        Assert.True(view.Type("5"));
        Assert.True(view.IsEditing);
        Assert.Equal("5", view.EditText);
        view.BoxKey(Key.Enter, ModifierKeys.None);

        Assert.False(view.IsEditing);
        Assert.Equal(50.0, sheet["B1"].Value.AsNumber());
        Assert.Equal(Cell("A2"), view.ActiveCell);
    });

    [Fact]
    public void Escape_leaves_the_cell_as_it_was() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);

        view.Type("9");
        view.BoxKey(Key.Escape, ModifierKeys.None);

        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
        Assert.False(view.IsEditing);
    });

    [Fact]
    public void F2_edits_the_formula() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.ActiveCell = Cell("B1");

        view.PressKey(Key.F2, ModifierKeys.None);
        Assert.Equal("=A1*10", view.EditText);
        Assert.False(view.BoxKey(Key.Left, ModifierKeys.None));   // the caret moves, the cell stays
        view.EditText = "=A1*100";
        view.BoxKey(Key.Tab, ModifierKeys.None);

        Assert.Equal(200.0, sheet["B1"].Value.AsNumber());
        Assert.Equal(Cell("C1"), view.ActiveCell);
    });

    [Fact]
    public void Ctrl_Enter_fills_the_selection_and_a_bad_formula_stays_open() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.Selection = GridSelection.Cell(Cell("C1"), sheet).ExtendTo(Cell("D2"), sheet);

        view.Type("7");
        view.BoxKey(Key.Enter, ModifierKeys.Control);
        Assert.Equal(7.0, sheet["D2"].Value.AsNumber());

        view.Type("=1+");
        view.BoxKey(Key.Enter, ModifierKeys.None);
        Assert.True(view.IsEditing);
        view.CancelEdit();
    });

    [Fact]
    public void Delete_undo_redo_and_paste() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.Selection = GridSelection.Cell(Cell("A1"), sheet).ExtendTo(Cell("B1"), sheet);

        view.PressKey(Key.Delete, ModifierKeys.None);
        Assert.Equal(CellValue.Empty, sheet["A1"].Value);

        ApplicationCommands.Undo.Execute(null, view);
        Assert.Equal(20.0, sheet["B1"].Value.AsNumber());
        ApplicationCommands.Redo.Execute(null, view);
        Assert.Equal(CellValue.Empty, sheet["A1"].Value);

        view.ActiveCell = Cell("C2");
        Clipboard.SetText("1\t2\r\n3\t=C2+C3\r\n");
        ApplicationCommands.Paste.Execute(null, view);
        Assert.Equal(4.0, sheet["D3"].Value.AsNumber());
        Assert.Equal("C2:D3", view.Selection?.ToString());
    });

    [Fact]
    public void A_read_only_view_changes_nothing() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.IsReadOnly = true;

        Assert.False(view.Type("5"));
        view.PressKey(Key.Delete, ModifierKeys.None);

        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
    });

    [Fact]
    public void The_formula_bar_shows_and_edits_the_active_cell() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        var bar = new FormulaBar { View = view };

        view.ActiveCell = Cell("B1");
        Assert.Equal("B1", bar.NameBox.Text);
        Assert.Equal("=A1*10", bar.TextBox.Text);

        view.ActiveCell = Cell("C1");
        view.BeginEdit(null, enterMode: false, focus: false);
        bar.TextBox.Text = "=A1+1";
        Assert.Equal("=A1+1", view.EditText);
        bar.TextKey(Key.Enter, ModifierKeys.None);

        Assert.Equal(3.0, sheet["C1"].Value.AsNumber());
        Assert.Equal(Cell("C2"), view.ActiveCell);

        bar.NameBox.Text = "D40:E41";
        bar.NameKey(Key.Enter);
        Assert.Equal("D40:E41", view.Selection?.ToString());
    });

    [Fact]
    public void Ctrl_arrows_and_home_move_as_in_Excel() => Sta.Run(() =>
    {
        var sheet = Sheet();
        sheet["A5"].Value = 1;
        var view = Show(sheet);

        view.PressKey(Key.Down, ModifierKeys.Control);
        Assert.Equal(Cell("A5"), view.ActiveCell);
        view.PressKey(Key.End, ModifierKeys.Control);
        Assert.Equal(Cell("B5"), view.ActiveCell);
        view.PressKey(Key.Home, ModifierKeys.None);
        Assert.Equal(Cell("A5"), view.ActiveCell);
    });

    [Fact]
    public void Cut_and_paste_move_the_cells() => Sta.Run(() =>
    {
        var sheet = Sheet();
        var view = Show(sheet);
        view.Selection = GridSelection.Cell(Cell("A1"), sheet).ExtendTo(Cell("B1"), sheet);

        ApplicationCommands.Cut.Execute(null, view);
        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
        view.ActiveCell = Cell("C3");
        ApplicationCommands.Paste.Execute(null, view);

        Assert.Equal(CellValue.Empty, sheet["A1"].Value);
        Assert.Equal("=A1*10", sheet["D3"].Formula);
    });
}
