using System;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.VisualTree;

namespace SharpCell.Grid.Avalonia;

public partial class SpreadsheetView
{
    /// <summary>Whether cells cannot be changed: no typing, Delete, pasting, undo or redo.</summary>
    public static readonly StyledProperty<bool> IsReadOnlyProperty =
        AvaloniaProperty.Register<SpreadsheetView, bool>(nameof(IsReadOnly));

    /// <summary>Whether a cell is being edited.</summary>
    public static readonly DirectProperty<SpreadsheetView, bool> IsEditingProperty =
        AvaloniaProperty.RegisterDirect<SpreadsheetView, bool>(nameof(IsEditing), v => v.IsEditing);

    /// <summary>The text being typed into the active cell; null when no cell is being edited.</summary>
    public static readonly DirectProperty<SpreadsheetView, string?> EditTextProperty =
        AvaloniaProperty.RegisterDirect<SpreadsheetView, string?>(nameof(EditText), v => v.EditText, (v, text) => v.EditText = text);

    private TextBox? _box;
    private bool _editing;
    private string? _editText;
    private bool _syncing;

    // The cell being edited: the active cell when editing began.
    private CellPosition _editCell;

    // Started by typing, the arrows enter the text and move as in Excel; started with F2, a
    // double-click or the formula bar, they move the caret.
    private bool _enterMode;

    /// <summary>Whether cells cannot be changed: no typing, Delete, pasting, undo or redo.</summary>
    public bool IsReadOnly
    {
        get => GetValue(IsReadOnlyProperty);
        set => SetValue(IsReadOnlyProperty, value);
    }

    /// <summary>Whether a cell is being edited.</summary>
    public bool IsEditing => _editing;

    /// <summary>
    /// The text being typed into the active cell, which <see cref="CommitEdit"/> enters; null when
    /// no cell is being edited. Setting it while editing changes the text, as a formula bar does.
    /// </summary>
    public string? EditText
    {
        get => _editText;
        set
        {
            if (!_editing || value is null || value == _editText)
                return;
            SetAndRaise(EditTextProperty, ref _editText, value);
            if (_box is not null && _box.Text != value)
            {
                _syncing = true;
                _box.Text = value;
                _syncing = false;
            }

            InvalidateArrange();
        }
    }

    /// <summary>
    /// What changes the cells, with the history to undo and redo; one per workbook, created when
    /// <see cref="Worksheet"/> is set. Null without a sheet.
    /// </summary>
    public GridEditor? Editor { get; private set; }

    /// <summary>After each change to the cells through the view or its <see cref="Editor"/>, undo and redo included.</summary>
    public event EventHandler<GridEdit>? Edited;

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
    /// Pastes the text on the clipboard at the selection's top-left cell, as Ctrl+V does: text
    /// copied from Excel goes in cell by cell, each read as typed. After <see cref="CutAsync"/>, the
    /// cut cells move there instead, formulas and formats with them. The pasted cells are selected.
    /// </summary>
    public async Task PasteAsync()
    {
        if (IsReadOnly || Worksheet is null || Selection is null || TopLevel.GetTopLevel(this)?.Clipboard is not { } clipboard)
            return;
        if (await clipboard.TryGetTextAsync() is not { } text || Worksheet is not { } sheet || Selection is not { } selection || Editor is not { } editor)
            return;
        var at = new CellPosition(selection.Area.FirstRow, selection.Area.FirstColumn);

        // Cut cells move once, while the clipboard still holds what was cut from them.
        if (_cut is { } cut && cut.Sheet == sheet && cut.Text == text)
        {
            _cut = null;
            _moving = true;
            try
            {
                Show(Try(() => editor.Move(sheet, cut.Area, at)));
            }
            finally
            {
                _moving = false;
            }

            return;
        }

        Show(Try(() => editor.Paste(sheet, at, text)));
    }

    /// <summary>
    /// Cuts the selected cells, as Ctrl+X does in Excel: they are copied, and the next paste moves
    /// them, formulas and formats with them; until then they stay. Escape or another change forgets
    /// the cut.
    /// </summary>
    public async Task CutAsync()
    {
        if (IsReadOnly || Worksheet is not { } sheet || Selection is not { } selection)
            return;
        await CopyAsync();
        _cut = new Cut(sheet, selection.Area, GridClipboard.Text(sheet, selection.Area));
    }

    // Cells cut and waiting for a paste, with the text the clipboard got for them.
    private sealed record Cut(Worksheet Sheet, CellArea Area, string Text);

    private Cut? _cut;
    private bool _moving;
    private bool _commandKey;

    /// <inheritdoc />
    protected override Size MeasureOverride(Size availableSize)
    {
        // The editor's size is its own: the view asks for no size, as without it.
        _box?.Measure(Size.Infinity);
        return default;
    }

    /// <inheritdoc />
    protected override Size ArrangeOverride(Size finalSize)
    {
        if (_box is not null)
            _box.Arrange(_editing ? EditorBounds(finalSize) : default);
        return finalSize;
    }

    /// <inheritdoc />
    protected override void OnTextInput(TextInputEventArgs e)
    {
        base.OnTextInput(e);

        // Typing into the grid starts editing the active cell with what was typed, replacing its
        // content; a letter typed with Ctrl or Cmd was a shortcut, not text.
        var shortcut = _commandKey;
        _commandKey = false;
        if (e.Handled || _editing || shortcut || string.IsNullOrEmpty(e.Text) || char.IsControl(e.Text[0]))
            return;
        if (BeginEdit(e.Text, enterMode: true, focus: true))
            e.Handled = true;
    }

    // The keys that change cells when nothing is being edited.
    private bool EditingKey(KeyEventArgs e, PlatformHotkeyConfiguration? hotkeys, bool command)
    {
        if (IsReadOnly)
            return false;
        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        if (hotkeys?.Paste.Exists(g => g.Matches(e)) == true || (command && e.Key == Key.V))
            _ = PasteAsync();
        else if (hotkeys?.Cut.Exists(g => g.Matches(e)) == true || (command && e.Key == Key.X))
            _ = CutAsync();
        else if (hotkeys?.Redo.Exists(g => g.Matches(e)) == true || (command && (e.Key == Key.Y || (e.Key == Key.Z && shift))))
            Redo();
        else if (hotkeys?.Undo.Exists(g => g.Matches(e)) == true || (command && e.Key == Key.Z))
            Undo();
        else if (e.Key == Key.F2 && e.KeyModifiers == KeyModifiers.None)
            BeginEdit();
        else if (e.Key == Key.Delete && e.KeyModifiers == KeyModifiers.None)
            ClearSelection();
        else if (e.Key == Key.Back && e.KeyModifiers == KeyModifiers.None)
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
        _syncing = true;
        box.Text = text;
        _syncing = false;
        box.CaretIndex = text.Length;
        box.IsVisible = true;
        ToolTip.SetTip(box, null);
        SetAndRaise(EditTextProperty, ref _editText, text);
        var was = _editing;
        _editing = true;
        RaisePropertyChanged(IsEditingProperty, was, true);
        ScrollTo(Painter.ScrollIntoView(Viewport, active));
        InvalidateMeasure();
        InvalidateArrange();
        if (focus)
            box.Focus();
        return true;
    }

    // Enters the text into the active cell, or with Ctrl+Enter into every selected cell.
    private bool Commit(bool all, Worksheet? sheetEdited = null)
    {
        if (!_editing)
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
        ToolTip.SetTip(_box, reason);
        ToolTip.SetIsOpen(_box, true);
        if (position is { } at)
            _box.CaretIndex = Math.Clamp(at, 0, _box.Text?.Length ?? 0);
        _enterMode = false;
        _box.Focus();
    }

    private void EndEdit()
    {
        if (!_editing)
            return;
        _editing = false;
        if (_box is not null)
        {
            ToolTip.SetIsOpen(_box, false);
            _box.IsVisible = false;
        }

        SetAndRaise(EditTextProperty, ref _editText, null);
        RaisePropertyChanged(IsEditingProperty, true, false);
        InvalidateArrange();
        InvalidateVisual();
        Focus();
    }

    // The keys of Excel while typing into a cell, before the text box sees them.
    private void OnBoxKeyDown(object? sender, KeyEventArgs e)
    {
        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        var alt = e.KeyModifiers.HasFlag(KeyModifiers.Alt);
        var control = e.KeyModifiers.HasFlag(KeyModifiers.Control) || e.KeyModifiers.HasFlag(KeyModifiers.Meta);
        switch (e.Key)
        {
            case Key.Enter when alt:
                // A line break inside the cell.
                _box!.SelectedText = "\n";
                break;
            case Key.Enter when control:
                Commit(all: true);
                break;
            case Key.Enter:
                CommitAndMove(selection => selection.MoveWithin(shift ? -1 : 1, 0, Worksheet));
                break;
            case Key.Tab:
                CommitAndMove(selection => selection.MoveWithin(0, shift ? -1 : 1, Worksheet));
                break;
            case Key.Escape:
                CancelEdit();
                break;
            case Key.F2:
                _enterMode = !_enterMode;
                break;
            case Key.Left or Key.Right or Key.Up or Key.Down when _enterMode && Step(e.Key) is var (rows, columns):
                CommitAndMove(selection => selection.Move(rows, columns, Worksheet));
                break;
            default:
                if (!CommandKeys.TextBoxKey(_box!, e))
                    return;
                break;
        }

        e.Handled = true;
    }

    // Enter and Tab from the formula bar: in, and on to the next cell.
    internal void CommitAndMove(int rows, int columns) =>
        CommitAndMove(selection => selection.MoveWithin(rows, columns, Worksheet));

    private void CommitAndMove(Func<GridSelection, GridSelection> move)
    {
        if (Commit(all: false) && Selection is { } selection)
            Go(move(selection), extent: false);
    }

    private void OnWorksheetChanged(Worksheet? old)
    {
        // What was typed goes into the sheet it was typed on; a formula that does not parse is dropped.
        if (_editing && (old is null || !Commit(all: false, old)))
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
    }

    private void OnEdited(object? sender, GridEdit edit)
    {
        // Any other change, typing or undo, ends the cut, as in Excel.
        if (!_moving)
            _cut = null;
        InvalidateVisual();
        Edited?.Invoke(this, edit);
    }

    // Selects the cells a change changed, as Excel does after undo, redo and paste.
    private void Show(GridEdit? edit)
    {
        if (edit is null || edit.Sheet != Worksheet)
            return;
        var area = edit.Area;
        Go(GridSelection.Cell(new CellPosition(area.FirstRow, area.FirstColumn), Worksheet).ExtendTo(new CellPosition(area.LastRow, area.LastColumn), Worksheet), extent: false);
    }

    private GridEdit? Try(Func<GridEdit> change)
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

    private bool IsInEditor(object? source) => _box is not null && source is Visual visual && (visual == _box || _box.IsVisualAncestorOf(visual));

    private TextBox Box()
    {
        if (_box is not null)
            return _box;
        _box = new TextBox
        {
            AcceptsReturn = false,
            MinHeight = 0,
            MinWidth = 0,
            Padding = new Thickness(1, 0, 1, 1),
            BorderThickness = new Thickness(0),
            CornerRadius = new CornerRadius(0),
            VerticalContentAlignment = VerticalAlignment.Bottom,
            IsVisible = false,
        };

        // Without the theme's focus border and backgrounds, the text box looks like the cell.
        foreach (var key in new[] { "TextControlBackground", "TextControlBackgroundPointerOver", "TextControlBackgroundFocused" })
            _box.Resources[key] = Brushes.White;
        foreach (var key in new[] { "TextControlBorderBrush", "TextControlBorderBrushPointerOver", "TextControlBorderBrushFocused" })
            _box.Resources[key] = Brushes.Transparent;
        _box.Resources["TextControlBorderThemeThicknessFocused"] = new Thickness(0);
        _box.AddHandler(KeyDownEvent, OnBoxKeyDown, RoutingStrategies.Tunnel);
        // The property, not TextChanged: that comes later, through the dispatcher.
        _box.PropertyChanged += (_, e) =>
        {
            if (e.Property != TextBox.TextProperty || _syncing || !_editing)
                return;
            SetAndRaise(EditTextProperty, ref _editText, _box.Text ?? "");
            InvalidateMeasure();
            InvalidateArrange();
        };
        VisualChildren.Add(_box);
        LogicalChildren.Add(_box);
        return _box;
    }

    // The cell's font, colour and fill, at its alignment.
    private static void StyleLike(TextBox box, Cell cell)
    {
        var style = cell.Style;
        var font = GridPainter.FontOf(style.Font);
        box.FontFamily = new FontFamily(font.Family);
        box.FontSize = font.Size;
        box.FontWeight = font.Bold ? FontWeight.Bold : FontWeight.Normal;
        box.FontStyle = font.Italic ? FontStyle.Italic : FontStyle.Normal;
        box.Foreground = style.Font.Color is { } color ? new SolidColorBrush(Color.FromRgb(color.R, color.G, color.B)) : (IBrush)Brushes.Black;
        IBrush fill = style.Fill is { } f ? new SolidColorBrush(Color.FromRgb(f.R, f.G, f.B)) : Brushes.White;
        foreach (var key in new[] { "TextControlBackground", "TextControlBackgroundPointerOver", "TextControlBackgroundFocused" })
            box.Resources[key] = fill;
        box.Background = fill;
        box.TextWrapping = style.WrapText ? TextWrapping.Wrap : TextWrapping.NoWrap;
    }

    // Over the active cell inside its border, widened to the right for longer text as in Excel.
    private Rect EditorBounds(Size size)
    {
        var cell = GridSelection.Cell(_editCell, Worksheet).Area;
        var area = Painter.CellBounds(Viewport, new CellPosition(cell.FirstRow, cell.FirstColumn), new CellPosition(cell.LastRow, cell.LastColumn));
        var x = Math.Floor(area.X) + 1;
        var y = Math.Floor(area.Y) + 1;
        var wanted = _box?.DesiredSize ?? default;
        var width = Math.Min(Math.Max(area.Width - 3, wanted.Width), Math.Max(0, size.Width - x));
        // One line keeps to the cell, inside its border; line breaks and wrapped text grow downwards.
        var lines = (_editText?.Contains('\n', StringComparison.Ordinal) ?? false) || _box?.TextWrapping == TextWrapping.Wrap;
        var height = Math.Min(lines ? Math.Max(area.Height - 3, wanted.Height) : area.Height - 3, Math.Max(0, size.Height - y));
        return new Rect(x, y, Math.Max(0, width), Math.Max(0, height));
    }
}
