using System.Runtime.CompilerServices;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input.Platform;
using SharpCell.Grid.Avalonia;
using SharpCell.Grid.Tests.Controls;
using SharpCell.Xlsx;

namespace SharpCell.Grid.Tests.Docs;

/// <summary>The code the documentation shows, between snippet markers, run as tests.</summary>
public class GuideExamples(HeadlessSession session)
{
    private static string Budget([CallerFilePath] string here = "") =>
        Path.GetFullPath(Path.Combine(Path.GetDirectoryName(here)!, "..", "..", "..", "samples", "SharpCell.Grid.Sample", "budget.xlsx"));

    [Fact]
    public Task Show_a_workbook() => session.Run(() =>
    {
        var path = Budget();

        // snippet: show-workbook
        Workbook workbook = XlsxReader.Load(path);
        workbook.Recalculate();

        var view = new SpreadsheetView { Worksheet = workbook["Budget"] };
        // end-snippet

        var window = new Window { Width = 600, Height = 300, Content = view };
        window.Show();
        Assert.NotNull(window.CaptureRenderedFrame());
        Assert.Equal(new CellPosition(1, 1), view.ActiveCell);
    });

    [Fact]
    public Task Refresh_after_changes() => session.Run(() =>
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        var view = new SpreadsheetView { Worksheet = sheet };

        // snippet: refresh
        sheet["B2"].Value = 99;
        workbook.Recalculate();
        view.InvalidateVisual();     // values, formulas and styles show at the next paint

        sheet.SetColumnWidth(2, 20);
        sheet.FrozenRows = 1;
        view.ReloadLayout();         // sizes and frozen panes are taken again
        // end-snippet

        Assert.Equal(1, view.Painter.FrozenRows);
        Assert.True(view.Painter.Layout.ColumnWidth(2) > 64);
    });

    [Fact]
    public Task Editing() => session.Run(() =>
    {
        var sheet = new Workbook().AddSheet("Sheet1");
        sheet["B1"].Formula = "=A1*2";
        sheet.Workbook.Recalculate();
        var view = new SpreadsheetView { Worksheet = sheet };
        var window = new Window { Content = view };
        window.Show();
        IReadOnlyList<Cell> changed = [];

        // snippet: editing
        var bar = new FormulaBar { View = view };   // Excel's formula bar, as a control of its own
        view.Edited += (_, edit) =>                 // after each change, the workbook recalculated
            changed = edit.Recalculation.ChangedCells;
        view.IsReadOnly = true;                      // to only show and select
        // end-snippet

        view.IsReadOnly = false;
        view.BeginEdit("21");
        view.CommitEdit();
        Assert.Equal(42.0, sheet["B1"].Value.AsNumber());
        Assert.Contains(changed, c => c.Address == "B1");
        Assert.Same(view, bar.View);
    });

    [Fact]
    public Task Rows_columns_and_zoom() => session.Run(() =>
    {
        var sheet = new Workbook().AddSheet("Sheet1");
        sheet["A1"].Value = 1;
        sheet["A2"].Value = 2;
        var view = new SpreadsheetView { Worksheet = sheet };
        var window = new Window { Content = view };
        window.Show();
        view.ActiveCell = new CellPosition(2, 1);

        // snippet: sheet-commands
        view.Zoom = 1.5;          // 150%, as Ctrl or Cmd with the wheel
        view.Editor!.Fill(sheet, new CellArea(1, 1, 2, 1), new CellArea(1, 1, 5, 1));   // as the fill handle: 1 to 5
        view.InsertRows();        // a row before the selected one, A2, as Ctrl+Plus
        view.FillDown();          // the row above copied into it, as Ctrl+D
        // end-snippet

        Assert.Equal(1.5, view.Painter.Zoom);
        Assert.Equal([1.0, 1.0, 2.0, 3.0, 4.0, 5.0], Enumerable.Range(1, 6).Select(r => sheet[r, 1].Value.AsNumber()));
    });

    [Fact]
    public Task Sheets_and_formatting() => session.Run(() =>
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sales");
        workbook.AddSheet("Costs");
        sheet["A1"].Value = 0.25;
        var view = new SpreadsheetView { Worksheet = sheet };
        var window = new Window { Content = view };
        window.Show();

        // snippet: sheets-and-formatting
        var tabs = new SheetTabs { View = view };   // Excel's sheet tabs, under the view
        view.ToggleBold();                          // as Ctrl+B
        view.SetFill(CellColor.FromRgb(0xFFF2CC));  // as Fill Color
        view.SetNumberFormat("0.0%");               // as the Number Format box
        view.ShowNextSheet();                       // as Ctrl+Page Down: the next sheet
        // end-snippet

        Assert.True(sheet["A1"].Style.Font.Bold);
        Assert.Equal("25.0%", sheet["A1"].Text);
        Assert.Equal(["Sales", "Costs"], tabs.SheetNames);
        Assert.Same(workbook["Costs"], view.Worksheet);
    });

    [Fact]
    public Task Selection_copying_and_scrolling() => session.Run(async () =>
    {
        var sheet = new Workbook().AddSheet("Sheet1");
        var view = new SpreadsheetView { Worksheet = sheet };
        var window = new Window { Content = view };
        window.Show();

        // snippet: active-cell
        view.ActiveCell = new CellPosition(row: 10, column: 3);                // C10, selected alone
        view.Selection = view.Selection!.Extend(rows: 2, columns: 1, sheet);   // C10:D12, as Shift with the arrows
        view.Offset = new Vector(0, 200);                                      // scrolled ten default rows down
        await view.CopyAsync();                                                // as Ctrl+C
        // end-snippet

        Assert.Equal("C10", view.ActiveCell.ToString());
        Assert.Equal("C10:D12", view.Selection.ToString());
        Assert.Equal("\t\r\n\t\r\n\t\r\n", await window.Clipboard!.TryGetTextAsync());
    });

    [Fact]
    public void Paint_without_a_UI_framework()
    {
        var workbook = new Workbook();
        var sheet = workbook.AddSheet("Sheet1");
        sheet["A1"].Value = "Total";
        sheet["B1"].Value = 1250;
        var canvas = new Painting.RecordingCanvas();

        // snippet: painter
        var painter = new GridPainter
        {
            // Text width in your UI framework's font engine, for fitting numbers and running text.
            MeasureText = (text, font) => text.Length * font.Size * 0.5,
        };
        painter.Layout = SheetLayout.FromWorksheet(sheet, painter.MaxDigitWidth(sheet));
        painter.FrozenRows = sheet.FrozenRows;
        painter.FrozenColumns = sheet.FrozenColumns;

        var viewport = new GridViewport(OffsetX: 0, OffsetY: 0, Width: 800, Height: 600);
        var selection = GridSelection.Cell(new CellPosition(1, 1), sheet);
        painter.Paint(canvas, sheet, viewport, selection);   // canvas: your IGridCanvas

        // A click, then a drag to (300, 90): what Excel selects, and what is copied.
        if (painter.Hit(viewport, x: 120, y: 30, sheet) is { } hit)
            selection = GridSelection.At(hit, sheet).ExtendTo(painter.CellNearest(viewport, 300, 90), sheet);
        string copied = GridClipboard.Text(sheet, selection.Area);
        // end-snippet

        Assert.Contains(canvas.Texts, t => t.Value == "Total");
        Assert.Equal("B1:E4", selection.ToString());
        Assert.StartsWith("1250\t\t\t\r\n", copied);
    }
}
