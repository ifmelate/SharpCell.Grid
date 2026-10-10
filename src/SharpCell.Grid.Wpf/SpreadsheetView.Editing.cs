using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace SharpCell.Grid.Wpf;

public partial class SpreadsheetView
{
    /// <summary>Whether cells cannot be changed: no typing, Delete, pasting, undo or redo.</summary>
    public static readonly DependencyProperty IsReadOnlyProperty = DependencyProperty.Register(
        nameof(IsReadOnly), typeof(bool), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(false, (d, _) => ((SpreadsheetView)d).Changed?.Invoke(d, EventArgs.Empty)));

    private static readonly DependencyPropertyKey IsEditingKey = DependencyProperty.RegisterReadOnly(
        nameof(IsEditing), typeof(bool), typeof(SpreadsheetView), new FrameworkPropertyMetadata(false));

    /// <summary>Whether a cell is being edited.</summary>
    public static readonly DependencyProperty IsEditingProperty = IsEditingKey.DependencyProperty;

    /// <summary>The text being typed into the active cell; null when no cell is being edited.</summary>
    public static readonly DependencyProperty EditTextProperty = DependencyProperty.Register(
        nameof(EditText), typeof(string), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(null, (d, e) => ((SpreadsheetView)d).OnEditTextChanged((string?)e.NewValue), (d, value) => ((SpreadsheetView)d).IsEditing ? value : null));

    private TextBox? _box;
    private bool _syncing;

    // The cell being edited: the active cell when editing began.
    private CellPosition _editCell;

    // Started by typing, the arrows enter the text and move as in Excel; started with F2, a
    // double-click or the formula bar, they move the caret.
    private bool _enterMode;

    /// <summary>Whether cells cannot be changed: no typing, Delete, pasting, undo or redo.</summary>
    public bool IsReadOnly
    {
        get => (bool)GetValue(IsReadOnlyProperty);
        set => SetValue(IsReadOnlyProperty, value);
    }

    /// <summary>Whether a cell is being edited.</summary>
    public bool IsEditing => (bool)GetValue(IsEditingProperty);

    /// <summary>
    /// The text being typed into the active cell, which <see cref="CommitEdit"/> enters; null when
    /// no cell is being edited. Setting it while editing changes the text, as a formula bar does.
    /// </summary>
    public string? EditText
    {
        get => (string?)GetValue(EditTextProperty);
        set => SetValue(EditTextProperty, value);
    }

    /// <summary>
    /// What changes the cells, with the history to undo and redo; one per workbook, created when
    /// <see cref="Worksheet"/> is set. Null without a sheet.
    /// </summary>
    public GridEditor? Editor { get; private set; }

    /// <summary>After each change to the cells through the view or its <see cref="Editor"/>, undo and redo included.</summary>
    public event EventHandler<GridEdit>? Edited;

    // What a formula bar shows changed: the selection, the sheet, editing, the text, read-only.
    internal event EventHandler? Changed;

    /// <inheritdoc />
    protected override int VisualChildrenCount => _box is null ? 0 : 1;

    /// <summary>
    /// Starts editing the active cell, as F2 does: with its formula or value as Excel shows it to
    /// edit, or with <paramref name="text"/>. False when the view is read-only or has no sheet.
    /// </summary>
    public bool BeginEdit(string? text = null) => BeginEdit(text, enterMode: false, focus: true);

    /// <summary>
    /// Enters the text being typed into the active cell, as Enter does, and ends editing; true
    /// also when nothing is being edited. A formula that does not parse keeps the editor open,
    /// with the reason shown and the caret where the formula went wrong, and gives false.
    /// </summary>
    public bool CommitEdit() => Commit(all: false);

    /// <summary>Ends editing without changing the cell, as Escape does.</summary>
    public void CancelEdit() => EndEdit();

    /// <summary>Undoes the last change to the cells and selects them, as Ctrl+Z does.</summary>
    public void Undo() => Show(IsReadOnly ? null : Editor?.Undo());

    /// <summary>Redoes the last undone change and selects its cells, as Ctrl+Y does.</summary>
    public void Redo() => Show(IsReadOnly ? null : Editor?.Redo());

    /// <summary>Clears the content of the selected cells, as Delete does; formats and styles stay.</summary>
    public void ClearSelection()
    {
        if (!IsReadOnly && Worksheet is { } sheet && Selection is { } selection && Editor is { } editor)
            Try(() => editor.Clear(sheet, selection.Area));
    }

    /// <summary>
    /// Pastes the text on the clipboard at the selection's top-left cell, as Ctrl+V and
    /// <see cref="ApplicationCommands.Paste"/> do: text copied from Excel goes in cell by cell, each
    /// read as typed. The pasted cells are selected.
    /// </summary>
    public void Paste()
    {
        if (IsReadOnly || Worksheet is not { } sheet || Selection is not { } selection || Editor is not { } editor || !Clipboard.ContainsText())
            return;
        var text = Clipboard.GetText();
        var at = new CellPosition(selection.Area.FirstRow, selection.Area.FirstColumn);
        Show(Try(() => editor.Paste(sheet, at, text)));
    }

    /// <inheritdoc />
    protected override Visual GetVisualChild(int index) =>
        index == 0 && _box is not null ? _box : throw new ArgumentOutOfRangeException(nameof(index));

    /// <inheritdoc />
    protected override Size MeasureOverride(Size availableSize)
    {
        _box?.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        return base.MeasureOverride(availableSize);
    }

    /// <inheritdoc />
    protected override Size ArrangeOverride(Size finalSize)
    {
        _box?.Arrange(IsEditing ? EditorBounds(finalSize) : default);
        return finalSize;
    }

    /// <inheritdoc />
    protected override void OnTextInput(TextCompositionEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnTextInput(e);
        if (!e.Handled && Type(e.Text))
            e.Handled = true;
    }

    // Typing into the grid starts editing the active cell with what was typed, replacing its content.
    internal bool Type(string text) =>
        !IsEditing && !string.IsNullOrEmpty(text) && !char.IsControl(text[0]) && BeginEdit(text, enterMode: true, focus: true);

    // The keys that change cells when nothing is being edited; paste, undo and redo are commands.
    private bool EditingKey(Key key, ModifierKeys modifiers)
    {
        if (IsReadOnly || modifiers != ModifierKeys.None)
            return false;
        if (key == Key.F2)
            BeginEdit();
        else if (key == Key.Delete)
            ClearSelection();
        else if (key == Key.Back)
            BeginEdit("", enterMode: true, focus: true);
        else
            return false;
        return true;
    }

    internal bool BeginEdit(string? text, bool enterMode, bool focus)
    {
        if (IsReadOnly || Worksheet is not { } sheet || Selection is not { } selection)
            return false;
        var active = selection.ActiveCell;
        _editCell = active;
        var cell = sheet[active.Row, active.Column];
        text ??= cell.Input;
        var box = Box();
        StyleLike(box, cell);
        _enterMode = enterMode;
        box.ToolTip = null;
        box.Visibility = Visibility.Visible;
        SetValue(IsEditingKey, true);
        SetText(text);
        box.CaretIndex = text.Length;
        ScrollTo(Painter.ScrollIntoView(Viewport, active));
        InvalidateMeasure();
        InvalidateArrange();
        if (focus)
            box.Focus();
        Changed?.Invoke(this, EventArgs.Empty);
        return true;
    }

    // Enters the text into the active cell, or with Ctrl+Enter into every selected cell.
    private bool Commit(bool all, Worksheet? sheetEdited = null)
    {
        if (!IsEditing)
            return true;
        if ((sheetEdited ?? Worksheet) is not { } sheet || Selection is not { } selection || Editor is not { } editor)
        {
            EndEdit();
            return true;
        }

        var active = _editCell;
        var area = all ? selection.Area : new CellArea(active.Row, active.Column, active.Row, active.Column);
        try
        {
            editor.Enter(sheet, area, _box?.Text ?? "");
        }
        catch (FormulaParseException error)
        {
            Refuse(error.Message, error.Position);
            return false;
        }
        catch (InvalidOperationException error)
        {
            Refuse(error.Message, null);
            return false;
        }

        EndEdit();
        return true;
    }

    private void Refuse(string reason, int? position)
    {
        if (_box is null)
            return;
        _box.ToolTip = new ToolTip { Content = reason, PlacementTarget = _box, IsOpen = true };
        if (position is { } at)
            _box.CaretIndex = Math.Clamp(at, 0, _box.Text.Length);
        _enterMode = false;
        _box.Focus();
    }

    private void EndEdit()
    {
        if (!IsEditing)
            return;
        if (_box is not null)
        {
            if (_box.ToolTip is ToolTip tip)
                tip.IsOpen = false;
            _box.ToolTip = null;
            _box.Visibility = Visibility.Collapsed;
        }

        SetValue(IsEditingKey, false);
        SetCurrentValue(EditTextProperty, null);
        InvalidateArrange();
        InvalidateVisual();
        Focus();
        Changed?.Invoke(this, EventArgs.Empty);
    }

    // The keys of Excel while typing into a cell, before the text box sees them.
    internal bool BoxKey(Key key, ModifierKeys modifiers)
    {
        var shift = modifiers.HasFlag(ModifierKeys.Shift);
        var alt = modifiers.HasFlag(ModifierKeys.Alt);
        var control = modifiers.HasFlag(ModifierKeys.Control);
        switch (key)
        {
            case Key.Enter when alt:
                // A line break inside the cell.
                _box!.SelectedText = "\n";
                _box.CaretIndex = _box.SelectionStart + 1;
                break;
            case Key.Enter when control:
                Commit(all: true);
                break;
            case Key.Enter:
                CommitAndMove(shift ? -1 : 1, 0);
                break;
            case Key.Tab:
                CommitAndMove(0, shift ? -1 : 1);
                break;
            case Key.Escape:
                CancelEdit();
                break;
            case Key.F2:
                _enterMode = !_enterMode;
                break;
            case Key.Left or Key.Right or Key.Up or Key.Down when _enterMode && Step(key) is var (rows, columns):
                if (Commit(all: false) && Selection is { } selection)
                    Go(selection.Move(rows, columns, Worksheet), extent: false);
                break;
            default:
                return false;
        }

        return true;
    }

    // Enter and Tab: in, and on to the next cell.
    internal void CommitAndMove(int rows, int columns)
    {
        if (Commit(all: false) && Selection is { } selection)
            Go(selection.MoveWithin(rows, columns, Worksheet), extent: false);
    }

    private void OnEditedSheetChanged(Worksheet? old)
    {
        // What was typed goes into the sheet it was typed on; a formula that does not parse is dropped.
        if (IsEditing && (old is null || !Commit(all: false, old)))
            EndEdit();

        if (Worksheet?.Workbook is not { } workbook)
        {
            Editor = null;
        }
        else if (Editor?.Workbook != workbook)
        {
            if (Editor is not null)
                Editor.Edited -= OnEdited;
            Editor = new GridEditor(workbook);
            Editor.Edited += OnEdited;
        }

        Changed?.Invoke(this, EventArgs.Empty);
    }

    private void OnEdited(object? sender, GridEdit edit)
    {
        InvalidateVisual();
        Edited?.Invoke(this, edit);
        Changed?.Invoke(this, EventArgs.Empty);
    }

    private void OnEditTextChanged(string? text)
    {
        if (text is not null && _box is not null && _box.Text != text)
        {
            _syncing = true;
            _box.Text = text;
            _syncing = false;
        }

        InvalidateMeasure();
        InvalidateArrange();
        Changed?.Invoke(this, EventArgs.Empty);
    }

    private void SetText(string text)
    {
        SetCurrentValue(EditTextProperty, text);
        OnEditTextChanged(text);
    }

    // Selects the cells a change changed, as Excel does after undo, redo and paste.
    private void Show(GridEdit? edit)
    {
        if (edit is null || edit.Sheet != Worksheet)
            return;
        var area = edit.Area;
        Go(GridSelection.Cell(new CellPosition(area.FirstRow, area.FirstColumn), Worksheet).ExtendTo(new CellPosition(area.LastRow, area.LastColumn), Worksheet), extent: false);
    }

    private static GridEdit? Try(Func<GridEdit> change)
    {
        try
        {
            return change();
        }
        catch (InvalidOperationException)
        {
            // Part of an array formula: Excel refuses the change, and so does the view.
            return null;
        }
    }

    private void CanChange(object sender, CanExecuteRoutedEventArgs e) => e.CanExecute = !IsReadOnly && Worksheet is not null && Selection is not null;

    private bool IsInEditor(object? source) => _box is not null && source is DependencyObject element && (element == _box || _box.IsAncestorOf(element));

    private TextBox Box()
    {
        if (_box is not null)
            return _box;
        _box = new TextBox
        {
            AcceptsReturn = false,
            AcceptsTab = false,
            BorderThickness = new Thickness(0),
            Padding = new Thickness(0, 0, 0, 1),
            VerticalContentAlignment = VerticalAlignment.Bottom,
            Visibility = Visibility.Collapsed,
        };
        _box.PreviewKeyDown += (_, e) =>
        {
            if (BoxKey(e.Key == Key.System ? e.SystemKey : e.Key, Keyboard.Modifiers))
                e.Handled = true;
        };
        _box.TextChanged += (_, _) =>
        {
            if (!_syncing && IsEditing)
                SetCurrentValue(EditTextProperty, _box.Text);
        };
        AddVisualChild(_box);
        AddLogicalChild(_box);
        return _box;
    }

    // The cell's font, colour and fill, at its alignment.
    private static void StyleLike(TextBox box, Cell cell)
    {
        var style = cell.Style;
        var font = GridPainter.FontOf(style.Font);
        box.FontFamily = new FontFamily(font.Family);
        box.FontSize = font.Size;
        box.FontWeight = font.Bold ? FontWeights.Bold : FontWeights.Normal;
        box.FontStyle = font.Italic ? FontStyles.Italic : FontStyles.Normal;
        box.Foreground = style.Font.Color is { } color ? new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B)) : Brushes.Black;
        box.Background = style.Fill is { } fill ? new SolidColorBrush(Color.FromRgb(fill.R, fill.G, fill.B)) : Brushes.White;
        box.TextWrapping = style.WrapText ? TextWrapping.Wrap : TextWrapping.NoWrap;
    }

    // Over the active cell inside its border, widened to the right for longer text as in Excel.
    private Rect EditorBounds(Size size)
    {
        if (_box is null)
            return default;
        var cell = GridSelection.Cell(_editCell, Worksheet).Area;
        var area = Painter.CellBounds(Viewport, new CellPosition(cell.FirstRow, cell.FirstColumn), new CellPosition(cell.LastRow, cell.LastColumn));
        var x = Math.Floor(area.X) + 1;
        var y = Math.Floor(area.Y) + 1;
        var wanted = _box.DesiredSize;
        var width = Math.Min(Math.Max(area.Width - 3, wanted.Width), Math.Max(0, size.Width - x));

        // One line keeps to the cell, inside its border; line breaks and wrapped text grow downwards.
        var lines = _box.Text.Contains('\n', StringComparison.Ordinal) || _box.TextWrapping == TextWrapping.Wrap;
        var height = Math.Min(lines ? Math.Max(area.Height - 3, wanted.Height) : area.Height - 3, Math.Max(0, size.Height - y));
        return new Rect(x, y, Math.Max(0, width), Math.Max(0, height));
    }
}
