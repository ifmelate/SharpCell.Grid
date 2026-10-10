using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace SharpCell.Grid.Wpf;

/// <summary>
/// Excel's formula bar for a <see cref="SpreadsheetView"/>: the name box with the active cell's
/// address, and the active cell's formula or value as Excel shows it to edit. Typing in the bar
/// edits the cell, and the cell's editor shows the same text; Enter and Tab enter it, Escape
/// leaves the cell as it was. An address or a range typed into the name box selects it.
/// </summary>
public class FormulaBar : UserControl
{
    /// <summary>The view whose active cell the bar shows and edits.</summary>
    public static readonly DependencyProperty ViewProperty = DependencyProperty.Register(
        nameof(View), typeof(SpreadsheetView), typeof(FormulaBar),
        new FrameworkPropertyMetadata(null, (d, e) => ((FormulaBar)d).OnViewChanged((SpreadsheetView?)e.OldValue)));

    private readonly TextBox _name;
    private readonly TextBox _text;
    private bool _syncing;

    /// <summary>A bar for no view; set <see cref="View"/> to connect it.</summary>
    public FormulaBar()
    {
        _name = new TextBox { Width = 90, VerticalContentAlignment = VerticalAlignment.Center };
        _text = new TextBox { AcceptsReturn = false, VerticalContentAlignment = VerticalAlignment.Center };
        var fx = new TextBlock { Text = "fx", FontStyle = FontStyles.Italic, Margin = new Thickness(8, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center, Foreground = Brushes.Gray };
        DockPanel.SetDock(_name, Dock.Left);
        DockPanel.SetDock(fx, Dock.Left);
        var panel = new DockPanel();
        panel.Children.Add(_name);
        panel.Children.Add(fx);
        panel.Children.Add(_text);
        Content = panel;

        _text.GotKeyboardFocus += (_, _) =>
        {
            if (View is { IsEditing: false } view)
                view.BeginEdit(null, enterMode: false, focus: false);
        };
        _text.TextChanged += (_, _) =>
        {
            if (!_syncing && View is { IsEditing: true } view)
                view.EditText = _text.Text;
        };
        _text.PreviewKeyDown += (_, e) => e.Handled = TextKey(e.Key == Key.System ? e.SystemKey : e.Key, Keyboard.Modifiers);
        _name.PreviewKeyDown += (_, e) => e.Handled = NameKey(e.Key);
    }

    /// <summary>The view whose active cell the bar shows and edits.</summary>
    public SpreadsheetView? View
    {
        get => (SpreadsheetView?)GetValue(ViewProperty);
        set => SetValue(ViewProperty, value);
    }

    internal TextBox NameBox => _name;

    internal TextBox TextBox => _text;

    private void OnViewChanged(SpreadsheetView? old)
    {
        if (old is not null)
            old.Changed -= OnChanged;
        if (View is { } view)
            view.Changed += OnChanged;
        Refresh();
    }

    private void OnChanged(object? sender, EventArgs e) => Refresh();

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

    internal bool TextKey(Key key, ModifierKeys modifiers)
    {
        if (View is not { IsEditing: true } view)
            return false;
        var shift = modifiers.HasFlag(ModifierKeys.Shift);
        switch (key)
        {
            case Key.Enter when modifiers.HasFlag(ModifierKeys.Alt):
                _text.SelectedText = "\n";
                _text.CaretIndex = _text.SelectionStart + 1;
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
                return false;
        }

        return true;
    }

    // An address such as B3 or a range such as B3:D7 selects it.
    internal bool NameKey(Key key)
    {
        if (View is not { } view || view.Worksheet is not { } sheet)
            return false;
        if (key == Key.Escape)
        {
            Refresh();
            view.Focus();
            return true;
        }

        if (key != Key.Enter)
            return false;
        try
        {
            var range = sheet.Range(_name.Text.Trim());
            view.Go(GridSelection.Cell(new CellPosition(range.FirstRow, range.FirstColumn), sheet)
                .ExtendTo(new CellPosition(range.LastRow, range.LastColumn), sheet), extent: false);
            view.Focus();
        }
        catch (Exception error) when (error is ArgumentException or FormatException)
        {
            Refresh();
        }

        return true;
    }
}
