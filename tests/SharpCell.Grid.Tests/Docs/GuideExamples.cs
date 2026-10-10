using System.Runtime.CompilerServices;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
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
    public Task Active_cell_and_scrolling() => session.Run(() =>
    {
        var view = new SpreadsheetView { Worksheet = new Workbook().AddSheet("Sheet1") };

        // snippet: active-cell
        view.ActiveCell = new CellPosition(row: 10, column: 3);   // C10
        view.Offset = new Vector(0, 200);                // scrolled ten default rows down
        // end-snippet

        Assert.Equal("C10", view.ActiveCell.ToString());
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
        painter.Paint(canvas, sheet, viewport, activeCell: new CellPosition(1, 1));   // canvas: your IGridCanvas
        CellPosition? clicked = painter.HitTest(viewport, x: 120, y: 30, sheet);
        // end-snippet

        Assert.Contains(canvas.Texts, t => t.Value == "Total");
        Assert.NotNull(clicked);
    }
}
