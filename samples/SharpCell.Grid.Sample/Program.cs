using System;
using System.IO;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Themes.Fluent;
using SharpCell;
using SharpCell.Grid.Avalonia;
using SharpCell.Xlsx;

internal static class Program
{
    [STAThread]
    public static void Main(string[] args) => BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);

    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<SampleApp>().UsePlatformDetect();
}

internal sealed class SampleApp : Application
{
    public override void Initialize() => Styles.Add(new FluentTheme());

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var path = desktop.Args is [var given, ..] ? given : Path.Combine(AppContext.BaseDirectory, "budget.xlsx");
            Workbook workbook = XlsxReader.Load(path);
            workbook.Recalculate();
            desktop.MainWindow = new Window
            {
                Title = $"{Path.GetFileName(path)} - SharpCell.Grid",
                Width = 960,
                Height = 600,
                Content = WithFormulaBar(new SpreadsheetView { Worksheet = workbook.Sheets[0] }),
            };
        }
        base.OnFrameworkInitializationCompleted();
    }

    private static DockPanel WithFormulaBar(SpreadsheetView view)
    {
        var bar = new FormulaBar { View = view, Margin = new Thickness(6) };
        DockPanel.SetDock(bar, Dock.Top);
        return new DockPanel { Children = { bar, view } };
    }
}
