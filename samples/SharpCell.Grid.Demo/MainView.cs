using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Platform.Storage;
using SharpCell.Grid.Avalonia;
using SharpCell.Xlsx;

namespace SharpCell.Grid.Demo;

/// <summary>A toolbar to open a file, the grid below it, and the sheet tabs at the bottom.</summary>
internal sealed class MainView : DockPanel
{
    private static readonly FilePickerFileType Xlsx = new("Excel workbook")
    {
        Patterns = ["*.xlsx"],
        MimeTypes = ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
    };

    private readonly SpreadsheetView _view = new();
    private readonly SheetTabs _tabs = new();
    private readonly TextBlock _status = new() { VerticalAlignment = VerticalAlignment.Center, TextWrapping = TextWrapping.Wrap };

    public MainView()
    {
        var open = new Button { Content = "Open an .xlsx file…", VerticalAlignment = VerticalAlignment.Center };
        open.Click += async (_, _) => await OpenAsync();
        var docs = new HyperlinkButton
        {
            Content = "Documentation",
            NavigateUri = new Uri("https://ifmelate.github.io/SharpCell.Grid/"),
            VerticalAlignment = VerticalAlignment.Center,
        };

        var bar = new DockPanel { Margin = new Thickness(8, 6), LastChildFill = true };
        var left = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 8, Children = { open } };
        DockPanel.SetDock(left, Dock.Left);
        DockPanel.SetDock(docs, Dock.Right);
        _status.Margin = new Thickness(12, 0);
        bar.Children.Add(left);
        bar.Children.Add(docs);
        bar.Children.Add(_status);

        DockPanel.SetDock(bar, Dock.Top);
        Children.Add(bar);
        var formulas = new FormulaBar { View = _view, Margin = new Thickness(8, 0, 8, 6) };
        DockPanel.SetDock(formulas, Dock.Top);
        Children.Add(formulas);
        _tabs.View = _view;
        DockPanel.SetDock(_tabs, Dock.Bottom);
        Children.Add(_tabs);
        Children.Add(_view);

        Show(Showcase.Build(), "An example built in code: type into the cells, the totals follow. Open your own file: it is read in your browser and never uploaded.");
    }

    private async Task OpenAsync()
    {
        if (TopLevel.GetTopLevel(this) is not { } top)
            return;
        var files = await top.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Open an .xlsx file",
            AllowMultiple = false,
            FileTypeFilter = [Xlsx],
        });
        if (files.Count == 0)
            return;

        var file = files[0];
        _status.Text = $"Reading {file.Name}…";
        try
        {
            // The browser's file stream is read asynchronously; the reader then works in memory.
            using var memory = new MemoryStream();
            await using (var stream = await file.OpenReadAsync())
                await stream.CopyToAsync(memory);
            memory.Position = 0;

            var watch = Stopwatch.StartNew();
            var workbook = XlsxReader.Load(memory);
            workbook.Recalculate();
            Show(workbook, $"{file.Name}: read and calculated by SharpCell in {watch.ElapsedMilliseconds} ms, in your browser.");
        }
        catch (Exception ex) when (ex is InvalidDataException or NotSupportedException)
        {
            _status.Text = $"{file.Name} cannot be read: {ex.Message}";
        }
    }

    private void Show(Workbook workbook, string message)
    {
        _view.Worksheet = workbook.Sheets[0];
        _view.ActiveCell = new CellPosition(1, 1);
        _view.Offset = default;
        _view.Focus();
        _status.Text = message;
    }
}
