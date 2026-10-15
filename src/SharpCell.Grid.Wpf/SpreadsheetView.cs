using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace SharpCell.Grid.Wpf;

/// <summary>
/// Shows a SharpCell worksheet as Excel does: headers, gridlines, each cell's text as Excel formats
/// it in the cell's style, the sheet's column widths, row heights, merged cells, frozen panes and
/// gridlines. Cells are selected as in Excel: by clicking and dragging, by Shift with a click or
/// the arrows, whole columns and rows from their headers or with Ctrl+Space and Shift+Space, the
/// whole sheet from the corner or with Ctrl+A. Enter and Tab move inside a selection, the wheel
/// scrolls and with Ctrl zooms, and Ctrl+C puts the selected cells on the clipboard as Excel pastes
/// them. Columns and rows are resized by dragging the edges of their headers, and fitted by
/// double-clicking them; the fill handle fills cells, and the headers have Excel's menu to insert,
/// delete and hide them. Only the visible cells are read, so any sheet size works. The view reads cells when it paints; after
/// changing cells or recalculating, call <see cref="UIElement.InvalidateVisual"/>. Sizes and frozen
/// panes are taken when <see cref="Worksheet"/> is set; after changing them, call <see cref="ReloadLayout"/>.
/// </summary>
public partial class SpreadsheetView : FrameworkElement
{
    /// <summary>The sheet shown; null shows an empty grid.</summary>
    public static readonly DependencyProperty WorksheetProperty = DependencyProperty.Register(
        nameof(Worksheet), typeof(Worksheet), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender, (d, e) => ((SpreadsheetView)d).OnWorksheetChanged((Worksheet?)e.OldValue)));

    /// <summary>How far the cells are scrolled, in pixels from cell A1.</summary>
    public static readonly DependencyProperty OffsetProperty = DependencyProperty.Register(
        nameof(Offset), typeof(Vector), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(default(Vector), FrameworkPropertyMetadataOptions.AffectsRender | FrameworkPropertyMetadataOptions.AffectsArrange));

    /// <summary>The cell with the active-cell border; null shows none.</summary>
    public static readonly DependencyProperty ActiveCellProperty = DependencyProperty.Register(
        nameof(ActiveCell), typeof(CellPosition?), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(new CellPosition(1, 1), FrameworkPropertyMetadataOptions.AffectsRender, (d, _) => ((SpreadsheetView)d).OnActiveCellChanged()));

    /// <summary>The selected cells, the active cell among them; null selects none.</summary>
    public static readonly DependencyProperty SelectionProperty = DependencyProperty.Register(
        nameof(Selection), typeof(GridSelection), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(GridSelection.Cell(new CellPosition(1, 1), null), FrameworkPropertyMetadataOptions.AffectsRender, (d, _) => ((SpreadsheetView)d).OnSelectionChanged()));

    // The selection and scrolling of sheets shown before, for when they are shown again.
    private readonly System.Runtime.CompilerServices.ConditionalWeakTable<Worksheet, SheetState> _sheetStates = [];

    // Rows one notch of the wheel scrolls, as in Excel; a notch is 120 in WPF's wheel deltas.
    private const int WheelRows = 3;
    private const double WheelNotch = 120;

    static SpreadsheetView()
    {
        FocusableProperty.OverrideMetadata(typeof(SpreadsheetView), new FrameworkPropertyMetadata(true));
        ClipToBoundsProperty.OverrideMetadata(typeof(SpreadsheetView), new FrameworkPropertyMetadata(true));
    }

    /// <summary>Creates an empty view; set <see cref="Worksheet"/> to show a sheet.</summary>
    public SpreadsheetView()
    {
        Painter = new GridPainter { MeasureText = Measure, ShowsFillHandle = true };
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Copy, (_, e) => { Copy(); e.Handled = true; }, (_, e) => e.CanExecute = Worksheet is not null && Selection is not null));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.SelectAll, (_, e) => { SelectAll(); e.Handled = true; }));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Paste, (_, e) => { Paste(); e.Handled = true; }, CanChange));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Cut, (_, e) => { Cut(); e.Handled = true; }, CanChange));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Undo, (_, e) => { Undo(); e.Handled = true; }, (_, e) => e.CanExecute = !IsReadOnly && Editor?.CanUndo == true));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Redo, (_, e) => { Redo(); e.Handled = true; }, (_, e) => e.CanExecute = !IsReadOnly && Editor?.CanRedo == true));
    }

    /// <summary>The sheet shown; null shows an empty grid.</summary>
    public Worksheet? Worksheet
    {
        get => (Worksheet?)GetValue(WorksheetProperty);
        set => SetValue(WorksheetProperty, value);
    }

    /// <summary>How far the cells are scrolled, in pixels from cell A1; kept inside the sheet when painting.</summary>
    public Vector Offset
    {
        get => (Vector)GetValue(OffsetProperty);
        set => SetValue(OffsetProperty, value);
    }

    /// <summary>
    /// The cell with the active-cell border; null shows none. It is the active cell of
    /// <see cref="Selection"/>; setting it selects that cell alone.
    /// </summary>
    public CellPosition? ActiveCell
    {
        get => (CellPosition?)GetValue(ActiveCellProperty);
        set => SetValue(ActiveCellProperty, value);
    }

    /// <summary>The selected cells, the active cell among them; null selects none. A new sheet starts from its active cell alone.</summary>
    public GridSelection? Selection
    {
        get => (GridSelection?)GetValue(SelectionProperty);
        set => SetValue(SelectionProperty, value);
    }

    /// <summary>Geometry and colours: row heights, column widths, header sizes, theme.</summary>
    public GridPainter Painter { get; }

    /// <summary>What the view shows now: its size and its offset, kept inside the sheet.</summary>
    public GridViewport Viewport => Painter.Clamp(new GridViewport(Offset.X, Offset.Y, ActualWidth, ActualHeight));

    /// <summary>Takes the sheet's column widths, row heights, hidden rows and columns and frozen panes again.</summary>
    public void ReloadLayout()
    {
        var sheet = Worksheet;
        var zoom = Painter.Zoom;
        Painter.Layout = sheet is not null
            ? SheetLayout.FromWorksheet(sheet, Painter.MaxDigitWidth(sheet), zoom)
            : new SheetLayout(Math.Round(SheetLayout.ExcelColumnWidth * zoom), Math.Round(SheetLayout.ExcelRowHeight * zoom));
        Painter.FrozenRows = sheet?.FrozenRows ?? 0;
        Painter.FrozenColumns = sheet?.FrozenColumns ?? 0;
        InvalidateVisual();
    }

    /// <summary>
    /// Puts the selected cells on the clipboard as Excel pastes them: as text with tabs between
    /// cells, and as an HTML table with their styles. Ctrl+C and <see cref="ApplicationCommands.Copy"/> do it too.
    /// </summary>
    public void Copy()
    {
        if (Worksheet is not { } sheet || Selection is not { } selection)
            return;
        var data = new DataObject();
        data.SetText(GridClipboard.Text(sheet, selection.Area), TextDataFormat.UnicodeText);
        var html = GridClipboard.WindowsHtml(GridClipboard.Html(sheet, selection.Area));
        data.SetData(DataFormats.Html, new MemoryStream(Encoding.UTF8.GetBytes(html)));
        Clipboard.SetDataObject(data, copy: true);

        // Pasted here again, the cells are copied with their formulas, as in Excel.
        _copied = new Clipped(sheet, selection.Area, GridClipboard.Text(sheet, selection.Area));
        _cut = null;
        Painter.CopiedArea = selection.Area;
        InvalidateVisual();
    }

    /// <summary>Paints the visible part of the sheet through <see cref="Painter"/>.</summary>
    /// <param name="drawingContext">The drawing context WPF renders with.</param>
    protected override void OnRender(DrawingContext drawingContext)
    {
        ArgumentNullException.ThrowIfNull(drawingContext);
        Painter.Paint(new WpfGridCanvas(drawingContext, PixelsPerDip), Worksheet, Viewport, Selection);
    }

    /// <inheritdoc />
    protected override void OnKeyDown(KeyEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnKeyDown(e);
        if (!e.Handled && !IsEditing && PressKey(e.Key, Keyboard.Modifiers))
            e.Handled = true;
    }

    // What a key does, as in Excel; whether it did anything.
    internal bool PressKey(Key key, ModifierKeys modifiers)
    {
        if (Selection is not { } selection)
            return false;
        if (EditingKey(key, modifiers))
            return true;
        var sheet = Worksheet;
        var shift = modifiers.HasFlag(ModifierKeys.Shift);
        var control = modifiers.HasFlag(ModifierKeys.Control);
        var step = Step(key);
        var paging = key is Key.Home or Key.End or Key.PageUp or Key.PageDown;
        if (key is Key.PageUp or Key.PageDown && control && !shift)
        {
            ShowNextSheet(key == Key.PageDown ? 1 : -1);
            return true;
        }

        GridSelection? next = null;
        if (key == Key.Space && control)
            next = selection.EntireColumns(sheet);
        else if (key == Key.Space && shift)
            next = selection.EntireRows(sheet);
        else if (step is var (rows, columns))
            next = (control, shift) switch
            {
                (true, true) => selection.ExtendToEdge(rows, columns, sheet),
                (true, false) => selection.MoveToEdge(rows, columns, sheet),
                (false, true) => selection.Extend(rows, columns, sheet),
                _ => selection.Move(rows, columns, sheet),
            };
        else if (paging)
            next = Paging(key, selection, control, shift);
        else if (key == Key.Enter)
            next = selection.MoveWithin(shift ? -1 : 1, 0, sheet);
        else if (key == Key.Tab)
            next = selection.MoveWithin(0, shift ? -1 : 1, sheet);
        if (next is null)
            return false;

        Go(next, extent: shift && (step is not null || key is Key.PageUp or Key.PageDown));
        return true;
    }

    // Home: the row's first column; Ctrl+Home: A1; Ctrl+End: the last used cell; Page Up and Page
    // Down: a screen of rows, scrolled with it.
    private GridSelection? Paging(Key key, GridSelection selection, bool jump, bool shift)
    {
        var sheet = Worksheet;
        var active = selection.ActiveCell;
        CellPosition? target = key switch
        {
            Key.Home when jump => new CellPosition(1, 1),
            Key.Home => new CellPosition(active.Row, 1),
            Key.End when jump => sheet?.UsedRange is { } used ? new CellPosition(used.LastRow, used.LastColumn) : new CellPosition(1, 1),
            _ => null,
        };
        if (target is { } cell)
            return shift ? selection.ExtendTo(cell, sheet) : GridSelection.Cell(cell, sheet);
        if (key == Key.End)
            return null;

        var visible = Painter.VisibleArea(Viewport);
        var page = Math.Max(1, visible.LastRow - visible.FirstRow);
        var sign = key == Key.PageDown ? 1 : -1;
        var layout = Painter.Layout;
        var first = Math.Clamp(visible.FirstRow + sign * page, 1, SheetLayout.MaxRow);
        ScrollTo(Viewport with { OffsetY = Viewport.OffsetY + layout.RowTop(first) - layout.RowTop(visible.FirstRow) });
        return shift ? selection.Extend(sign * page, 0, sheet) : selection.Move(sign * page, 0, sheet);
    }

    // Selects and scrolls: to the moving corner when extending, to the active cell when moving.
    internal void Go(GridSelection next, bool extent)
    {
        Selection = next;
        if (!next.IsEntireColumns && !next.IsEntireRows || next.IsSingleCell)
            ScrollTo(Painter.ScrollIntoView(Viewport, extent ? next.Extent : next.ActiveCell));
    }

    private void SelectAll()
    {
        if (Selection is { } selection)
            Selection = selection.SelectAll(Worksheet);
    }

    private static (int Rows, int Columns)? Step(Key key) => key switch
    {
        Key.Left => (0, -1),
        Key.Right => (0, 1),
        Key.Up => (-1, 0),
        Key.Down => (1, 0),
        _ => null,
    };

    private void OnWorksheetChanged(Worksheet? old)
    {
        OnEditedSheetChanged(old);
        ReloadLayout();

        // Each sheet of a workbook keeps its selection and scrolling, as Excel's tabs do; a sheet
        // of the same workbook not shown before starts at A1.
        if (old is not null && Selection is { } leaving)
            _sheetStates.AddOrUpdate(old, new SheetState(leaving, Offset));
        if (Worksheet is { } shown && _sheetStates.TryGetValue(shown, out var state))
        {
            SetCurrentValue(SelectionProperty, state.Selection);
            SetCurrentValue(OffsetProperty, state.Offset);
        }
        else if (Worksheet is not null && old is not null && old.Workbook == Worksheet.Workbook)
        {
            SetCurrentValue(SelectionProperty, GridSelection.Cell(new CellPosition(1, 1), Worksheet));
            SetCurrentValue(OffsetProperty, default(Vector));
        }
        else if (ActiveCell is { } cell)
        {
            SetCurrentValue(SelectionProperty, GridSelection.Cell(cell, Worksheet));
        }

        Changed?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Shows the next sheet of the workbook, or with -1 the one before, as Ctrl+Page Down and
    /// Ctrl+Page Up do; false when there is none that way.
    /// </summary>
    public bool ShowNextSheet(int direction = 1)
    {
        if (Worksheet is not { } sheet)
            return false;
        var sheets = sheet.Workbook.Sheets;
        var index = -1;
        for (var i = 0; i < sheets.Count; i++)
        {
            if (sheets[i] == sheet)
                index = i;
        }

        var next = index + Math.Sign(direction);
        if (index < 0 || next < 0 || next >= sheets.Count)
            return false;
        Worksheet = sheets[next];
        return true;
    }

    private sealed record SheetState(GridSelection Selection, Vector Offset);

    private void OnSelectionChanged()
    {
        Changed?.Invoke(this, EventArgs.Empty);
        if (ActiveCell != Selection?.ActiveCell)
            SetCurrentValue(ActiveCellProperty, Selection?.ActiveCell);
    }

    private void OnActiveCellChanged()
    {
        if (Selection?.ActiveCell != ActiveCell)
            SetCurrentValue(SelectionProperty, ActiveCell is { } active ? GridSelection.Cell(active, Worksheet) : null);
    }

    private double PixelsPerDip => VisualTreeHelper.GetDpi(this).PixelsPerDip;

    private double Measure(string text, GridFont font) => WpfGridCanvas.Format(text, font, Brushes.Black, PixelsPerDip).Width;

    private void ScrollTo(GridViewport viewport)
    {
        var clamped = Painter.Clamp(viewport);
        Offset = new Vector(clamped.OffsetX, clamped.OffsetY);
    }
}
