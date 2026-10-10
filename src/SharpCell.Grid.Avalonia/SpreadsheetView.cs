using System;
using System.Text;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Media;

namespace SharpCell.Grid.Avalonia;

/// <summary>
/// Shows a SharpCell worksheet as Excel does: headers, gridlines, each cell's text as Excel formats
/// it in the cell's style, the sheet's column widths, row heights, merged cells, frozen panes and
/// gridlines. Cells are selected as in Excel: by clicking and dragging, by Shift with a click or
/// the arrows, whole columns and rows from their headers or with Ctrl+Space and Shift+Space, the
/// whole sheet from the corner or with Ctrl+A or Cmd+A. Enter and Tab move inside a selection, the
/// wheel scrolls, and Ctrl+C or Cmd+C puts the selected cells on the clipboard as Excel pastes them. Only the
/// visible cells are read, so any sheet size works. The view reads cells when it paints; after changing cells or
/// recalculating, call <see cref="Visual.InvalidateVisual"/>. Sizes and frozen panes are taken when
/// <see cref="Worksheet"/> is set; after changing them, call <see cref="ReloadLayout"/>.
/// </summary>
public class SpreadsheetView : Control
{
    /// <summary>The sheet shown; null shows an empty grid.</summary>
    public static readonly StyledProperty<Worksheet?> WorksheetProperty =
        AvaloniaProperty.Register<SpreadsheetView, Worksheet?>(nameof(Worksheet));

    /// <summary>How far the cells are scrolled, in pixels from cell A1.</summary>
    public static readonly StyledProperty<Vector> OffsetProperty =
        AvaloniaProperty.Register<SpreadsheetView, Vector>(nameof(Offset));

    /// <summary>The cell with the active-cell border; null shows none.</summary>
    public static readonly StyledProperty<CellPosition?> ActiveCellProperty =
        AvaloniaProperty.Register<SpreadsheetView, CellPosition?>(nameof(ActiveCell), new CellPosition(1, 1));

    /// <summary>The selected cells, the active cell among them; null selects none.</summary>
    public static readonly StyledProperty<GridSelection?> SelectionProperty =
        AvaloniaProperty.Register<SpreadsheetView, GridSelection?>(nameof(Selection), GridSelection.Cell(new CellPosition(1, 1), null));

    // Rows one turn of the wheel scrolls, as in Excel.
    private const int WheelRows = 3;

    static SpreadsheetView()
    {
        AffectsRender<SpreadsheetView>(WorksheetProperty, OffsetProperty, ActiveCellProperty, SelectionProperty);
        FocusableProperty.OverrideDefaultValue<SpreadsheetView>(true);
        ClipToBoundsProperty.OverrideDefaultValue<SpreadsheetView>(true);
    }

    /// <summary>The sheet shown; null shows an empty grid.</summary>
    public Worksheet? Worksheet
    {
        get => GetValue(WorksheetProperty);
        set => SetValue(WorksheetProperty, value);
    }

    /// <summary>How far the cells are scrolled, in pixels from cell A1; kept inside the sheet when painting.</summary>
    public Vector Offset
    {
        get => GetValue(OffsetProperty);
        set => SetValue(OffsetProperty, value);
    }

    /// <summary>
    /// The cell with the active-cell border; null shows none. It is the active cell of
    /// <see cref="Selection"/>; setting it selects that cell alone.
    /// </summary>
    public CellPosition? ActiveCell
    {
        get => GetValue(ActiveCellProperty);
        set => SetValue(ActiveCellProperty, value);
    }

    /// <summary>The selected cells, the active cell among them; null selects none. A new sheet starts from its active cell alone.</summary>
    public GridSelection? Selection
    {
        get => GetValue(SelectionProperty);
        set => SetValue(SelectionProperty, value);
    }

    /// <summary>Creates an empty view; set <see cref="Worksheet"/> to show a sheet.</summary>
    public SpreadsheetView()
    {
        Painter = new GridPainter { MeasureText = Measure };
    }

    /// <summary>Takes the sheet's column widths, row heights, hidden rows and columns and frozen panes again.</summary>
    public void ReloadLayout()
    {
        var sheet = Worksheet;
        Painter.Layout = sheet is not null ? SheetLayout.FromWorksheet(sheet, Painter.MaxDigitWidth(sheet)) : new SheetLayout();
        Painter.FrozenRows = sheet?.FrozenRows ?? 0;
        Painter.FrozenColumns = sheet?.FrozenColumns ?? 0;
        InvalidateVisual();
    }

    /// <inheritdoc />
    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == WorksheetProperty)
        {
            ReloadLayout();
            if (ActiveCell is { } cell)
                SetCurrentValue(SelectionProperty, GridSelection.Cell(cell, Worksheet));
        }
        else if (change.Property == SelectionProperty && ActiveCell != Selection?.ActiveCell)
        {
            SetCurrentValue(ActiveCellProperty, Selection?.ActiveCell);
        }
        else if (change.Property == ActiveCellProperty && Selection?.ActiveCell != ActiveCell)
        {
            SetCurrentValue(SelectionProperty, ActiveCell is { } active ? GridSelection.Cell(active, Worksheet) : null);
        }
    }

    /// <summary>Geometry and colours: row heights, column widths, header sizes, theme.</summary>
    public GridPainter Painter { get; }

    /// <summary>What the view shows now: its size and its offset, kept inside the sheet.</summary>
    public GridViewport Viewport => Painter.Clamp(new GridViewport(Offset.X, Offset.Y, Bounds.Width, Bounds.Height));

    /// <summary>Paints the visible part of the sheet through <see cref="Painter"/>.</summary>
    /// <param name="context">The drawing context Avalonia renders with.</param>
    public override void Render(DrawingContext context) =>
        Painter.Paint(new AvaloniaGridCanvas(context), Worksheet, Viewport, Selection);

    /// <summary>
    /// Puts the selected cells on the clipboard as Excel pastes them: as text with tabs between
    /// cells, and as an HTML table with their styles. Ctrl+C or Cmd+C does it too, also in a browser.
    /// </summary>
    public async Task CopyAsync()
    {
        if (Worksheet is not { } sheet || Selection is not { } selection || TopLevel.GetTopLevel(this)?.Clipboard is not { } clipboard)
            return;
        var html = GridClipboard.Html(sheet, selection.Area);
        var item = new DataTransferItem();
        item.SetText(GridClipboard.Text(sheet, selection.Area));

        // Each platform has its own name for HTML; Windows also wants it in its HTML Format.
        if (OperatingSystem.IsWindows())
            item.Set(DataFormat.CreateBytesPlatformFormat("HTML Format"), Encoding.UTF8.GetBytes(GridClipboard.WindowsHtml(html)));
        else
            item.Set(DataFormat.CreateStringPlatformFormat(OperatingSystem.IsMacOS() ? "public.html" : "text/html"), html);
        var data = new DataTransfer();
        data.Add(item);
        await clipboard.SetDataAsync(data);
    }

    private bool _dragging;

    /// <inheritdoc />
    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            return;
        Focus();
        e.Handled = true;
        var point = e.GetPosition(this);
        if (Painter.Hit(Viewport, point.X, point.Y, Worksheet) is not { } hit)
            return;
        Selection = e.KeyModifiers.HasFlag(KeyModifiers.Shift) && Selection is { } selection
            ? selection.ExtendTo(hit, Worksheet)
            : GridSelection.At(hit, Worksheet);
        _dragging = true;
        e.Pointer.Capture(this);
    }

    /// <inheritdoc />
    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (!_dragging || Selection is not { } selection)
            return;

        // Past an edge the nearest cell is the next one out of view, and scrolling to it follows the mouse.
        var point = e.GetPosition(this);
        var cell = Painter.CellNearest(Viewport, point.X, point.Y);
        Selection = selection.ExtendTo(cell, Worksheet);
        ScrollTo(Painter.ScrollIntoView(Viewport, cell));
        e.Handled = true;
    }

    /// <inheritdoc />
    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (!_dragging)
            return;
        _dragging = false;
        e.Pointer.Capture(null);
        e.Handled = true;
    }

    /// <inheritdoc />
    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e);
        _dragging = false;
    }

    /// <inheritdoc />
    protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
    {
        base.OnPointerWheelChanged(e);
        var layout = Painter.Layout;
        var step = new Vector(e.Delta.X * WheelRows * layout.DefaultColumnWidth, e.Delta.Y * WheelRows * layout.DefaultRowHeight);
        ScrollTo(Viewport with { OffsetX = Viewport.OffsetX - step.X, OffsetY = Viewport.OffsetY - step.Y });
        e.Handled = true;
    }

    /// <inheritdoc />
    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (Selection is not { } selection)
            return;
        var sheet = Worksheet;
        // The platform's gestures, and both Ctrl and Cmd with the letter: in a browser on macOS the
        // platform is the browser, whose gestures use Ctrl, while people press Cmd.
        var hotkeys = Application.Current?.PlatformSettings?.HotkeyConfiguration;
        var command = e.KeyModifiers is KeyModifiers.Control or KeyModifiers.Meta;
        if (hotkeys?.Copy.Exists(g => g.Matches(e)) == true || (command && e.Key is Key.C or Key.Insert))
        {
            _ = CopyAsync();
            e.Handled = true;
            return;
        }

        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        var control = e.KeyModifiers.HasFlag(KeyModifiers.Control);
        GridSelection? next = null;
        if (hotkeys?.SelectAll.Exists(g => g.Matches(e)) == true || (command && e.Key == Key.A))
            next = GridSelection.All(selection.ActiveCell);
        else if (e.Key == Key.Space && control)
            next = selection.EntireColumns(sheet);
        else if (e.Key == Key.Space && shift)
            next = selection.EntireRows(sheet);
        else if (Step(e.Key) is var (rows, columns))
            next = shift ? selection.Extend(rows, columns, sheet) : selection.Move(rows, columns, sheet);
        else if (e.Key == Key.Enter)
            next = selection.MoveWithin(shift ? -1 : 1, 0, sheet);
        else if (e.Key == Key.Tab)
            next = selection.MoveWithin(0, shift ? -1 : 1, sheet);
        if (next is null)
            return;

        // Extending scrolls to the moving corner, moving to the active cell.
        Selection = next;
        var shown = shift && Step(e.Key) is not null ? next.Extent : next.ActiveCell;
        if (!next.IsEntireColumns && !next.IsEntireRows || next.IsSingleCell)
            ScrollTo(Painter.ScrollIntoView(Viewport, shown));
        e.Handled = true;
    }

    private static (int Rows, int Columns)? Step(Key key) => key switch
    {
        Key.Left => (0, -1),
        Key.Right => (0, 1),
        Key.Up => (-1, 0),
        Key.Down => (1, 0),
        _ => null,
    };

    private static double Measure(string text, GridFont font) => AvaloniaGridCanvas.Format(text, font, null).Width;

    private void ScrollTo(GridViewport viewport)
    {
        var clamped = Painter.Clamp(viewport);
        Offset = new Vector(clamped.OffsetX, clamped.OffsetY);
    }
}
