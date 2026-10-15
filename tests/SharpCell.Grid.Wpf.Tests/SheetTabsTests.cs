using System.Windows.Input;

namespace SharpCell.Grid.Wpf.Tests;

public class SheetTabsTests
{
    private static Workbook Book()
    {
        var workbook = new Workbook();
        workbook.AddSheet("Sales")["A1"].Value = 1;
        workbook.AddSheet("Costs")["A1"].Value = 2;
        return workbook;
    }

    [Fact]
    public void Tabs_show_the_sheets_and_switch_the_view_which_keeps_each_selection() => Sta.Run(() =>
    {
        var workbook = Book();
        var view = new SpreadsheetView { Worksheet = workbook["Sales"] };
        var tabs = new SheetTabs { View = view };
        view.ActiveCell = new CellPosition(5, 3);

        Assert.Equal(["Sales", "Costs"], tabs.SheetNames);
        tabs.Show(workbook["Costs"]);
        Assert.Same(workbook["Costs"], view.Worksheet);
        Assert.Equal(new CellPosition(1, 1), view.ActiveCell);

        Assert.True(view.PressKey(Key.PageUp, ModifierKeys.Control));
        Assert.Same(workbook["Sales"], view.Worksheet);
        Assert.Equal(new CellPosition(5, 3), view.ActiveCell);

        Assert.Equal("Sheet3", tabs.AddSheet()!.Name);
        Assert.Equal(["Sales", "Costs", "Sheet3"], tabs.SheetNames);
    });
}
