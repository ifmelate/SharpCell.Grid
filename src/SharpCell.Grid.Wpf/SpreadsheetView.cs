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
/// scrolls, and Ctrl+C puts the selected cells on the clipboard as Excel pastes them. Only the
/// visible cells are read, so any sheet size works. The view reads cells when it paints; after
/// changing cells or recalculating, call <see cref="UIElement.InvalidateVisual"/>. Sizes and frozen
/// panes are taken when <see cref="Worksheet"/> is set; after changing them, call <see cref="ReloadLayout"/>.
/// </summary>
public class SpreadsheetView : FrameworkElement
{
    /// <summary>The sheet shown; null shows an empty grid.</summary>
    public static readonly DependencyProperty WorksheetProperty = DependencyProperty.Register(
        nameof(Worksheet), typeof(Worksheet), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.AffectsRender, (d, _) => ((SpreadsheetView)d).OnWorksheetChanged()));

    /// <summary>How far the cells are scrolled, in pixels from cell A1.</summary>
    public static readonly DependencyProperty OffsetProperty = DependencyProperty.Register(
        nameof(Offset), typeof(Vector), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(default(Vector), FrameworkPropertyMetadataOptions.AffectsRender));

    /// <summary>The cell with the active-cell border; null shows none.</summary>
    public static readonly DependencyProperty ActiveCellProperty = DependencyProperty.Register(
        nameof(ActiveCell), typeof(CellPosition?), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(new CellPosition(1, 1), FrameworkPropertyMetadataOptions.AffectsRender, (d, _) => ((SpreadsheetView)d).OnActiveCellChanged()));

    /// <summary>The selected cells, the active cell among them; null selects none.</summary>
    public static readonly DependencyProperty SelectionProperty = DependencyProperty.Register(
        nameof(Selection), typeof(GridSelection), typeof(SpreadsheetView),
        new FrameworkPropertyMetadata(GridSelection.Cell(new CellPosition(1, 1), null), FrameworkPropertyMetadataOptions.AffectsRender, (d, _) => ((SpreadsheetView)d).OnSelectionChanged()));

    // Rows one notch of the wheel scrolls, as in Excel; a notch is 120 in WPF's wheel deltas.
    private const int WheelRows = 3;
    private const double WheelNotch = 120;

    private bool _dragging;

    static SpreadsheetView()
    {
        FocusableProperty.OverrideMetadata(typeof(SpreadsheetView), new FrameworkPropertyMetadata(true));
        ClipToBoundsProperty.OverrideMetadata(typeof(SpreadsheetView), new FrameworkPropertyMetadata(true));
    }

    /// <summary>Creates an empty view; set <see cref="Worksheet"/> to show a sheet.</summary>
    public SpreadsheetView()
    {
        Painter = new GridPainter { MeasureText = Measure };
        CommandBindings.Add(new CommandBinding(ApplicationCommands.Copy, (_, e) => { Copy(); e.Handled = true; }, (_, e) => e.CanExecute = Worksheet is not null && Selection is not null));
        CommandBindings.Add(new CommandBinding(ApplicationCommands.SelectAll, (_, e) => { SelectAll(); e.Handled = true; }));
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
        Painter.Layout = sheet is not null ? SheetLayout.FromWorksheet(sheet, Painter.MaxDigitWidth(sheet)) : new SheetLayout();
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
    }

    /// <summary>Paints the visible part of the sheet through <see cref="Painter"/>.</summary>
    /// <param name="drawingContext">The drawing context WPF renders with.</param>
    protected override void OnRender(DrawingContext drawingContext)
    {
        ArgumentNullException.ThrowIfNull(drawingContext);
        Painter.Paint(new WpfGridCanvas(drawingContext, PixelsPerDip), Worksheet, Viewport, Selection);
    }

    /// <inheritdoc />
    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnMouseLeftButtonDown(e);
        Focus();
        e.Handled = true;
        if (Press(e.GetPosition(this), Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)))
            CaptureMouse();
    }

    /// <inheritdoc />
    protected override void OnMouseMove(MouseEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnMouseMove(e);
        if (_dragging && IsMouseCaptured)
        {
            DragTo(e.GetPosition(this));
            e.Handled = true;
        }
    }

    /// <inheritdoc />
    protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnMouseLeftButtonUp(e);
        if (!_dragging)
            return;
        _dragging = false;
        ReleaseMouseCapture();
        e.Handled = true;
    }

    /// <inheritdoc />
    protected override void OnLostMouseCapture(MouseEventArgs e)
    {
        base.OnLostMouseCapture(e);
        _dragging = false;
    }

    /// <inheritdoc />
    protected override void OnMouseWheel(MouseWheelEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnMouseWheel(e);
        var rows = e.Delta / WheelNotch * WheelRows;
        ScrollTo(Viewport with { OffsetY = Viewport.OffsetY - rows * Painter.Layout.DefaultRowHeight });
        e.Handled = true;
    }

    /// <inheritdoc />
    protected override void OnKeyDown(KeyEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnKeyDown(e);
        if (!e.Handled && PressKey(e.Key, Keyboard.Modifiers))
            e.Handled = true;
    }

    // What a press of the left button selects; whether it began a drag.
    internal bool Press(Point point, bool shift)
    {
        if (Painter.Hit(Viewport, point.X, point.Y, Worksheet) is not { } hit)
            return false;
        Selection = shift && Selection is { } selection ? selection.ExtendTo(hit, Worksheet) : GridSelection.At(hit, Worksheet);
        _dragging = true;
        return true;
    }

    // Past an edge the nearest cell is the next one out of view, and scrolling to it follows the mouse.
    internal void DragTo(Point point)
    {
        if (Selection is not { } selection)
            return;
        var cell = Painter.CellNearest(Viewport, point.X, point.Y);
        Selection = selection.ExtendTo(cell, Worksheet);
        ScrollTo(Painter.ScrollIntoView(Viewport, cell));
    }

    // What a key does, as in Excel; whether it did anything.
    internal bool PressKey(Key key, ModifierKeys modifiers)
    {
        if (Selection is not { } selection)
            return false;
        var sheet = Worksheet;
        var shift = modifiers.HasFlag(ModifierKeys.Shift);
        var control = modifiers.HasFlag(ModifierKeys.Control);
        var step = Step(key);
        GridSelection? next = null;
        if (key == Key.Space && control)
            next = selection.EntireColumns(sheet);
        else if (key == Key.Space && shift)
            next = selection.EntireRows(sheet);
        else if (step is var (rows, columns))
            next = shift ? selection.Extend(rows, columns, sheet) : selection.Move(rows, columns, sheet);
        else if (key == Key.Enter)
            next = selection.MoveWithin(shift ? -1 : 1, 0, sheet);
        else if (key == Key.Tab)
            next = selection.MoveWithin(0, shift ? -1 : 1, sheet);
        if (next is null)
            return false;

        // Extending scrolls to the moving corner, moving to the active cell.
        Selection = next;
        var shown = shift && step is not null ? next.Extent : next.ActiveCell;
        if (!next.IsEntireColumns && !next.IsEntireRows || next.IsSingleCell)
            ScrollTo(Painter.ScrollIntoView(Viewport, shown));
        return true;
    }

    private void SelectAll()
    {
        if (Selection is { } selection)
            Selection = GridSelection.All(selection.ActiveCell);
    }

    private static (int Rows, int Columns)? Step(Key key) => key switch
    {
        Key.Left => (0, -1),
        Key.Right => (0, 1),
        Key.Up => (-1, 0),
        Key.Down => (1, 0),
        _ => null,
    };

    private void OnWorksheetChanged()
    {
        ReloadLayout();
        if (ActiveCell is { } cell)
            SetCurrentValue(SelectionProperty, GridSelection.Cell(cell, Worksheet));
    }

    private void OnSelectionChanged()
    {
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
