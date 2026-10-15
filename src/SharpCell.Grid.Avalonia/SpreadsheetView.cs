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
public partial class SpreadsheetView : Control
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

    // The selection and scrolling of sheets shown before, for when they are shown again.
    private readonly System.Runtime.CompilerServices.ConditionalWeakTable<Worksheet, SheetState> _sheetStates = [];

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
        Painter = new GridPainter { MeasureText = Measure, ShowsFillHandle = true };
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

    /// <inheritdoc />
    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == WorksheetProperty)
        {
            var old = change.GetOldValue<Worksheet?>();
            OnWorksheetChanged(old);
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
        }
        else if (change.Property == OffsetProperty || change.Property == BoundsProperty)
        {
            InvalidateArrange();
        }
        else if (change.Property == ZoomProperty)
        {
            OnZoomChanged();
        }
        else if (change.Property == IsReadOnlyProperty)
        {
            Painter.ShowsFillHandle = !IsReadOnly;
            InvalidateVisual();
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

        // Pasted here again, the cells are copied with their formulas, as in Excel.
        _copied = new Clipped(sheet, selection.Area, GridClipboard.Text(sheet, selection.Area));
        _cut = null;
        Painter.CopiedArea = selection.Area;
        InvalidateVisual();
    }

    /// <inheritdoc />
    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (IsEditing || e.Handled || Selection is not { } selection)
            return;
        var sheet = Worksheet;

        // The platform's gestures, and Ctrl and Cmd alike for every shortcut (see CommandKeys).
        var hotkeys = Application.Current?.PlatformSettings?.HotkeyConfiguration;
        var command = CommandKeys.IsCommand(e.KeyModifiers);
        _commandKey = command;
        if (hotkeys?.Copy.Exists(g => g.Matches(e)) == true || (command && e.Key is Key.C or Key.Insert))
        {
            _ = CopyAsync();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Escape && (_cut ?? _copied) is not null)
        {
            ForgetClip();
            e.Handled = true;
            return;
        }

        if (EditingKey(e, hotkeys, command))
        {
            e.Handled = true;
            return;
        }

        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        var jump = command;
        GridSelection? next = null;
        if (hotkeys?.SelectAll.Exists(g => g.Matches(e)) == true || (command && e.Key == Key.A) || (command && shift && e.Key == Key.Space))
            next = selection.SelectAll(sheet);
        else if (e.Key == Key.Space && command)
            next = selection.EntireColumns(sheet);
        else if (e.Key == Key.Space && shift)
            next = selection.EntireRows(sheet);
        else if (Step(e.Key) is var (rows, columns))
            next = (jump, shift) switch
            {
                (true, true) => selection.ExtendToEdge(rows, columns, sheet),
                (true, false) => selection.MoveToEdge(rows, columns, sheet),
                (false, true) => selection.Extend(rows, columns, sheet),
                _ => selection.Move(rows, columns, sheet),
            };
        else if (e.Key is Key.PageUp or Key.PageDown && jump && !shift)
        {
            ShowNextSheet(e.Key == Key.PageDown ? 1 : -1);
            e.Handled = true;
            return;
        }
        else if (e.Key is Key.Home or Key.End or Key.PageUp or Key.PageDown)
            next = Paging(e.Key, selection, jump, shift);
        else if (e.Key == Key.Enter)
            next = selection.MoveWithin(shift ? -1 : 1, 0, sheet);
        else if (e.Key == Key.Tab)
            next = selection.MoveWithin(0, shift ? -1 : 1, sheet);
        if (next is null)
            return;

        Go(next, extent: shift && (Step(e.Key) is not null || e.Key is Key.PageUp or Key.PageDown));
        e.Handled = true;
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

    private static (int Rows, int Columns)? Step(Key key) => key switch
    {
        Key.Left => (0, -1),
        Key.Right => (0, 1),
        Key.Up => (-1, 0),
        Key.Down => (1, 0),
        _ => null,
    };

    private sealed record SheetState(GridSelection Selection, Vector Offset);

    private static double Measure(string text, GridFont font) => AvaloniaGridCanvas.Format(text, font, null).Width;

    private void ScrollTo(GridViewport viewport)
    {
        var clamped = Painter.Clamp(viewport);
        Offset = new Vector(clamped.OffsetX, clamped.OffsetY);
    }
}
