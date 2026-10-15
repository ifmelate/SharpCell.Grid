using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using SharpCell.Grid.Avalonia;

namespace SharpCell.Grid.Tests.Controls;

public class SheetTabsTests(HeadlessSession session)
{
    private static Workbook Book()
    {
        var workbook = new Workbook();
        workbook.AddSheet("Sales")["A1"].Value = 1;
        workbook.AddSheet("Costs")["A1"].Value = 2;
        workbook.Recalculate();
        return workbook;
    }

    private static (Window Window, SpreadsheetView View, SheetTabs Tabs) Show(Workbook workbook)
    {
        var view = new SpreadsheetView { Worksheet = workbook.Sheets[0] };
        var tabs = new SheetTabs { View = view };
        DockPanel.SetDock(tabs, Dock.Bottom);
        var window = new Window { Width = 480, Height = 240, Content = new DockPanel { Children = { tabs, view } } };
        window.Show();
        view.Focus();
        return (window, view, tabs);
    }

    // The middle of the tab with a name, in the window.
    private static Point TabCenter(Window window, SheetTabs tabs, string name)
    {
        var text = tabs.GetVisualDescendants().OfType<TextBlock>().Single(t => t.Text == name);
        var tab = (Visual)text.GetVisualParent()!;
        return tab.TranslatePoint(new Point(tab.Bounds.Width / 2, tab.Bounds.Height / 2), window)!.Value;
    }

    [Fact]
    public Task Clicking_a_tab_shows_its_sheet_and_each_sheet_keeps_its_selection() => session.Run(() =>
    {
        var workbook = Book();
        var (window, view, tabs) = Show(workbook);
        view.ActiveCell = Cells.Parse("C5");

        Assert.Equal(["Sales", "Costs"], tabs.SheetNames);
        var costs = TabCenter(window, tabs, "Costs");
        window.MouseDown(costs, MouseButton.Left);
        window.MouseUp(costs, MouseButton.Left);

        Assert.Same(workbook["Costs"], view.Worksheet);
        Assert.Equal(Cells.Parse("A1"), view.ActiveCell);                 // not shown before: A1
        view.ActiveCell = Cells.Parse("B2");

        var sales = TabCenter(window, tabs, "Sales");
        window.MouseDown(sales, MouseButton.Left);
        window.MouseUp(sales, MouseButton.Left);
        Assert.Equal(Cells.Parse("C5"), view.ActiveCell);
        Assert.True(view.ShowNextSheet());
        Assert.Equal(Cells.Parse("B2"), view.ActiveCell);
    });

    [Fact]
    public Task The_plus_button_adds_a_sheet_and_shows_it() => session.Run(() =>
    {
        var workbook = Book();
        workbook.AddSheet("Sheet3");
        var (_, view, tabs) = Show(workbook);

        var added = tabs.AddSheet();

        Assert.Equal("Sheet4", added!.Name);                              // Sheet3 is taken
        Assert.Same(added, view.Worksheet);
        Assert.Equal(["Sales", "Costs", "Sheet3", "Sheet4"], tabs.SheetNames);

        view.IsReadOnly = true;
        Assert.Null(tabs.AddSheet());
    });

    [Fact]
    public Task Ctrl_Page_Down_and_Up_go_through_the_sheets() => session.Run(() =>
    {
        var workbook = Book();
        var (window, view, _) = Show(workbook);

        window.KeyPress(Key.PageDown, RawInputModifiers.Control, PhysicalKey.PageDown, null);
        Assert.Same(workbook["Costs"], view.Worksheet);
        window.KeyPress(Key.PageDown, RawInputModifiers.Control, PhysicalKey.PageDown, null);
        Assert.Same(workbook["Costs"], view.Worksheet);                   // the last sheet
        window.KeyPress(Key.PageUp, RawInputModifiers.Control, PhysicalKey.PageUp, null);
        Assert.Same(workbook["Sales"], view.Worksheet);
    });

    [Fact]
    public Task Renders_the_tabs_as_approved() => session.Run(() =>
    {
        var (window, _, _) = Show(Book());

        Snapshot.Match(window.CaptureRenderedFrame(), "sheet-tabs");
    });
}
