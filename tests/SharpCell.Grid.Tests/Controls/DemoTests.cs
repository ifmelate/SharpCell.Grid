using Avalonia.Controls;
using SharpCell.Grid.Avalonia;
using SharpCell.Grid.Demo;

namespace SharpCell.Grid.Tests.Controls;

/// <summary>The live demo's workbook in a view with a formula bar, as the demo shows it.</summary>
public class DemoTests(HeadlessSession session)
{
    // The demo once stopped at start: its title, four words, crashed reading text as a date.
    [Fact]
    public Task Every_cell_of_the_demo_shows_in_the_formula_bar() => session.Run(() =>
    {
        var workbook = Showcase.Build();
        var view = new SpreadsheetView();
        var bar = new FormulaBar { View = view };
        DockPanel.SetDock(bar, Dock.Top);
        var window = new Window { Width = 1000, Height = 600, Content = new DockPanel { Children = { bar, view } } };
        window.Show();

        foreach (var sheet in workbook.Sheets)
        {
            view.Worksheet = sheet;
            foreach (var cell in sheet.Cells.ToList())
                view.ActiveCell = new CellPosition(cell.Row, cell.Column);
        }

        Assert.Equal(workbook.Sheets[^1], view.Worksheet);
    });

    // A workbook's culture is invariant unless set, and Ctrl+Shift+$ then gave "900.00 ¤".
    [Fact]
    public void The_demo_formats_currency_in_dollars() =>
        Assert.Equal("$#,##0.00_);($#,##0.00)", GridEditor.ShortcutFormat('$', Showcase.Build().Culture));
}
