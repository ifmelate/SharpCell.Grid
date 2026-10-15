using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SharpCell.Grid.Wpf;

/// <summary>
/// Excel's sheet tabs for a <see cref="SpreadsheetView"/>: a tab for each sheet of the view's
/// workbook, the shown one highlighted. Clicking a tab shows its sheet, and the + button adds a
/// sheet, as in Excel. Put it under the view, as Excel does.
/// </summary>
public class SheetTabs : UserControl
{
    /// <summary>The view whose workbook's sheets the tabs show and switch.</summary>
    public static readonly DependencyProperty ViewProperty = DependencyProperty.Register(
        nameof(View), typeof(SpreadsheetView), typeof(SheetTabs),
        new FrameworkPropertyMetadata(null, (d, e) => ((SheetTabs)d).OnViewChanged((SpreadsheetView?)e.OldValue)));

    private static readonly Brush Accent = Frozen(new SolidColorBrush(Color.FromRgb(0x10, 0x7C, 0x41)));
    private static readonly Brush Text = Frozen(new SolidColorBrush(Color.FromRgb(0x44, 0x44, 0x44)));

    private readonly StackPanel _tabs;
    private readonly Border _add;
    private readonly List<(Worksheet Sheet, Border Tab)> _shown = [];
    private Worksheet? _current;

    /// <summary>Tabs for no view; set <see cref="View"/> to connect them.</summary>
    public SheetTabs()
    {
        _tabs = new StackPanel { Orientation = Orientation.Horizontal };
        _add = new Border
        {
            Child = new TextBlock { Text = "+", FontSize = 16, Foreground = Text, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center },
            Width = 28,
            Background = Brushes.Transparent,
            ToolTip = "New sheet",
        };
        _add.MouseLeftButtonDown += (_, e) =>
        {
            AddSheet();
            e.Handled = true;
        };
        Content = new Border
        {
            Background = Frozen(new SolidColorBrush(Color.FromRgb(0xF3, 0xF3, 0xF3))),
            BorderBrush = Frozen(new SolidColorBrush(Color.FromRgb(0xC8, 0xC8, 0xC8))),
            BorderThickness = new Thickness(0, 1, 0, 0),
            Height = 28,
            Child = new ScrollViewer
            {
                Content = _tabs,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
            },
        };
    }

    /// <summary>The view whose workbook's sheets the tabs show and switch.</summary>
    public SpreadsheetView? View
    {
        get => (SpreadsheetView?)GetValue(ViewProperty);
        set => SetValue(ViewProperty, value);
    }

    /// <summary>The names on the tabs, in the workbook's order.</summary>
    public IReadOnlyList<string> SheetNames => [.. _shown.Select(t => t.Sheet.Name)];

    /// <summary>
    /// Adds a sheet after the others, named as Excel names a new sheet (<c>Sheet2</c>, <c>Sheet3</c>,
    /// the first number free), and shows it. Null when the view is read-only or has no sheet.
    /// </summary>
    public Worksheet? AddSheet()
    {
        if (View is not { IsReadOnly: false, Worksheet: { } shown } view)
            return null;
        var workbook = shown.Workbook;
        var number = workbook.Sheets.Count + 1;
        while (workbook.TryGetSheet("Sheet" + number.ToString(CultureInfo.InvariantCulture), out _))
            number++;
        var sheet = workbook.AddSheet("Sheet" + number.ToString(CultureInfo.InvariantCulture));
        view.Worksheet = sheet;
        Refresh();
        return sheet;
    }

    /// <summary>Shows the tabs again from the workbook's sheets, after sheets were added in code.</summary>
    public void Refresh()
    {
        _tabs.Children.Clear();
        _shown.Clear();
        var view = View;
        _current = view?.Worksheet;
        _add.Visibility = view is { IsReadOnly: false, Worksheet: not null } ? Visibility.Visible : Visibility.Collapsed;
        if (_current is not { } current)
            return;
        foreach (var sheet in current.Workbook.Sheets)
        {
            var active = sheet == current;
            var tab = new Border
            {
                Padding = new Thickness(14, 0, 14, 0),
                Background = active ? Brushes.White : Brushes.Transparent,
                BorderBrush = Accent,
                BorderThickness = new Thickness(0, 0, 0, active ? 2 : 0),
                Child = new TextBlock
                {
                    Text = sheet.Name,
                    VerticalAlignment = VerticalAlignment.Center,
                    FontWeight = active ? FontWeights.SemiBold : FontWeights.Normal,
                    Foreground = active ? Accent : Text,
                },
            };
            var shown = sheet;
            tab.MouseLeftButtonDown += (_, e) =>
            {
                Show(shown);
                e.Handled = true;
            };
            _tabs.Children.Add(tab);
            _shown.Add((sheet, tab));
        }

        // The + button comes right after the last tab, as in Excel.
        _tabs.Children.Add(_add);
    }

    /// <summary>Shows a sheet in the view, as clicking its tab does.</summary>
    internal void Show(Worksheet sheet)
    {
        if (View is not { } view)
            return;
        view.Worksheet = sheet;
        view.Focus();
    }

    private void OnViewChanged(SpreadsheetView? old)
    {
        if (old is not null)
            old.Changed -= OnChanged;
        if (View is { } view)
            view.Changed += OnChanged;
        Refresh();
    }

    // The view tells of every change; the tabs are made again when its sheet or read-only state did.
    private void OnChanged(object? sender, EventArgs e)
    {
        if (View is { } view && (view.Worksheet != _current || (_add.Visibility == Visibility.Visible) == view.IsReadOnly))
            Refresh();
    }

    private static Brush Frozen(Brush brush)
    {
        brush.Freeze();
        return brush;
    }
}
