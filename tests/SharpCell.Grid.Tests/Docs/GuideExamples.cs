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
