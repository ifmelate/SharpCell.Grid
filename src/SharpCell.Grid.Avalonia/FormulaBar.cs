using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;

namespace SharpCell.Grid.Avalonia;

/// <summary>
/// Excel's formula bar for a <see cref="SpreadsheetView"/>: the name box with the active cell's
/// address, and the active cell's formula or value as Excel shows it to edit. Typing in the bar
/// edits the cell, and the cell's editor shows the same text; Enter and Tab enter it, Escape
/// leaves the cell as it was. An address or a range typed into the name box selects it.
/// </summary>
public class FormulaBar : UserControl
{
    /// <summary>The view whose active cell the bar shows and edits.</summary>
    public static readonly StyledProperty<SpreadsheetView?> ViewProperty =
        AvaloniaProperty.Register<FormulaBar, SpreadsheetView?>(nameof(View));

    private readonly TextBox _name;
    private readonly TextBox _text;
    private bool _syncing;

    /// <summary>A bar for no view; set <see cref="View"/> to connect it.</summary>
    public FormulaBar()
    {
        _name = new TextBox { Width = 90, VerticalContentAlignment = VerticalAlignment.Center, MinHeight = 0 };
        _text = new TextBox { AcceptsReturn = false, VerticalContentAlignment = VerticalAlignment.Center, MinHeight = 0 };
        var fx = new TextBlock { Text = "fx", FontStyle = FontStyle.Italic, Margin = new Thickness(8, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = Brushes.Gray };
        DockPanel.SetDock(_name, Dock.Left);
        DockPanel.SetDock(fx, Dock.Left);
        Content = new DockPanel { Children = { _name, fx, _text } };

        _text.GotFocus += (_, _) =>
        {
            if (View is { IsEditing: false } view)
                view.BeginEdit(null, enterMode: false, focus: false);
        };
        // The property, not TextChanged: that comes later, through the dispatcher.
        _text.PropertyChanged += (_, e) =>
        {
            if (e.Property == TextBox.TextProperty && !_syncing && View is { IsEditing: true } view)
                view.EditText = _text.Text ?? "";
        };
        _text.AddHandler(KeyDownEvent, OnTextKeyDown, RoutingStrategies.Tunnel);
        _name.AddHandler(KeyDownEvent, OnNameKeyDown, RoutingStrategies.Tunnel);
    }

    /// <summary>The view whose active cell the bar shows and edits.</summary>
    public SpreadsheetView? View
    {
        get => GetValue(ViewProperty);
        set => SetValue(ViewProperty, value);
    }

    /// <inheritdoc />
    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property != ViewProperty)
            return;
        if (change.GetOldValue<SpreadsheetView?>() is { } old)
        {
            old.PropertyChanged -= OnViewChanged;
            old.Edited -= OnEdited;
        }

        if (View is { } view)
        {
            view.PropertyChanged += OnViewChanged;
            view.Edited += OnEdited;
        }

        Refresh();
    }

    private void OnViewChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (e.Property == SpreadsheetView.SelectionProperty || e.Property == SpreadsheetView.WorksheetProperty
            || e.Property == SpreadsheetView.IsEditingProperty || e.Property == SpreadsheetView.EditTextProperty
            || e.Property == SpreadsheetView.IsReadOnlyProperty)
            Refresh();
    }

    private void OnEdited(object? sender, GridEdit e) => Refresh();

    // What the bar shows: the text being typed while editing, otherwise the active cell's.
    private void Refresh()
    {
        var view = View;
        var sheet = view?.Worksheet;
        var active = view?.ActiveCell;
        _name.Text = active?.ToString() ?? "";
        var text = view is { IsEditing: true } ? view.EditText ?? ""
            : sheet is not null && active is { } cell ? sheet[cell.Row, cell.Column].Input
            : "";
        if (_text.Text != text)
        {
            _syncing = true;
            _text.Text = text;
            _syncing = false;
        }

        _text.IsReadOnly = view is null || sheet is null || view.IsReadOnly;
    }

    private void OnTextKeyDown(object? sender, KeyEventArgs e)
    {
        if (View is not { IsEditing: true } view)
            return;
        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        switch (e.Key)
        {
            case Key.Enter when e.KeyModifiers.HasFlag(KeyModifiers.Alt):
                _text.SelectedText = "\n";
                break;
            case Key.Enter:
                view.CommitAndMove(shift ? -1 : 1, 0);
                break;
            case Key.Tab:
                view.CommitAndMove(0, shift ? -1 : 1);
                break;
            case Key.Escape:
                view.CancelEdit();
                break;
            default:
                return;
        }

        e.Handled = true;
    }

    // An address such as B3 or a range such as B3:D7 selects it.
    private void OnNameKeyDown(object? sender, KeyEventArgs e)
    {
        if (View is not { } view || view.Worksheet is not { } sheet)
            return;
        if (e.Key == Key.Escape)
        {
            Refresh();
            view.Focus();
            e.Handled = true;
            return;
        }

        if (e.Key != Key.Enter)
            return;
        e.Handled = true;
        try
        {
            var range = sheet.Range((_name.Text ?? "").Trim());
            view.Go(GridSelection.Cell(new CellPosition(range.FirstRow, range.FirstColumn), sheet)
                .ExtendTo(new CellPosition(range.LastRow, range.LastColumn), sheet), extent: false);
            view.Focus();
        }
        catch (Exception error) when (error is ArgumentException or FormatException)
        {
            Refresh();
        }
    }
}
