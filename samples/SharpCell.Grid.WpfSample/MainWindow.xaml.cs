using System;
using System.IO;
using System.Windows;
using SharpCell.Xlsx;

namespace SharpCell.Grid.WpfSample;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        var args = Environment.GetCommandLineArgs();
        var path = args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "budget.xlsx");
        Title = $"{Path.GetFileName(path)} - SharpCell.Grid";

        // snippet: show-workbook-wpf
        Workbook workbook = XlsxReader.Load(path);
        workbook.Recalculate();
        View.Worksheet = workbook.Sheets[0];   // View: <grid:SpreadsheetView x:Name="View" /> in MainWindow.xaml
        // end-snippet
    }
}
