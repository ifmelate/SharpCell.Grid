using System;
using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Media;

namespace SharpCell.Grid.Avalonia;

/// <summary>
/// Shows a SharpCell worksheet as Excel does: headers, gridlines and each cell's text as Excel
/// formats it. The mouse picks the active cell, the keyboard moves it as in Excel, the wheel
/// scrolls. Only the visible cells are read, so any sheet size works. The view reads the sheet
/// when it paints; after changing cells or recalculating, call <see cref="Visual.InvalidateVisual"/>.
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

    // Rows one turn of the wheel scrolls, as in Excel.
    private const int WheelRows = 3;

    static SpreadsheetView()
    {
        AffectsRender<SpreadsheetView>(WorksheetProperty, OffsetProperty, ActiveCellProperty);
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

    /// <summary>The cell with the active-cell border; null shows none.</summary>
    public CellPosition? ActiveCell
    {
        get => GetValue(ActiveCellProperty);
        set => SetValue(ActiveCellProperty, value);
    }

    /// <summary>Creates an empty view; set <see cref="Worksheet"/> to show a sheet.</summary>
    public SpreadsheetView()
    {
        Painter = new GridPainter { MeasureText = Measure };
    }

    /// <summary>Geometry and colours: row heights, column widths, header sizes, theme.</summary>
    public GridPainter Painter { get; }

    /// <summary>What the view shows now: its size and its offset, kept inside the sheet.</summary>
    public GridViewport Viewport => Painter.Clamp(new GridViewport(Offset.X, Offset.Y, Bounds.Width, Bounds.Height));

    /// <inheritdoc />
    public override void Render(DrawingContext context) =>
        Painter.Paint(new AvaloniaGridCanvas(context, Typeface, TextElement.GetFontSize(this)), Worksheet, Viewport, ActiveCell);

    /// <inheritdoc />
    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            return;
        var point = e.GetPosition(this);
        if (Painter.HitTest(Viewport, point.X, point.Y) is { } cell)
            ActiveCell = cell;
        Focus();
        e.Handled = true;
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
        if (ActiveCell is not { } active)
            return;

        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        (int Rows, int Columns)? move = e.Key switch
        {
            Key.Left => (0, -1),
            Key.Right => (0, 1),
            Key.Up => (-1, 0),
            Key.Down => (1, 0),
            Key.Enter => (shift ? -1 : 1, 0),
            Key.Tab => (0, shift ? -1 : 1),
            _ => null,
        };
        if (move is not var (rows, columns))
            return;

        var target = new CellPosition(
            Math.Clamp(active.Row + rows, 1, SheetLayout.MaxRow),
            Math.Clamp(active.Column + columns, 1, SheetLayout.MaxColumn));
        ActiveCell = target;
        ScrollTo(Painter.ScrollIntoView(Viewport, target));
        e.Handled = true;
    }

    private Typeface Typeface => new(TextElement.GetFontFamily(this));

    private double Measure(string text) =>
        new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, Typeface, TextElement.GetFontSize(this), null).Width;

    private void ScrollTo(GridViewport viewport)
    {
        var clamped = Painter.Clamp(viewport);
        Offset = new Vector(clamped.OffsetX, clamped.OffsetY);
    }
}
