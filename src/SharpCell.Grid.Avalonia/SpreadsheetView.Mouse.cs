using System;
using Avalonia;
using Avalonia.Input;

namespace SharpCell.Grid.Avalonia;

// The mouse as in Excel: selecting, dragging the edges of headers to resize columns and rows and
// double-clicking them to fit, dragging the fill handle, the wheel to scroll and with Ctrl to zoom.
public partial class SpreadsheetView
{
    private enum Drag
    {
        None,
        Select,
        ColumnEdge,
        RowEdge,
        Fill,
        Point,
    }

    // Made when first shown: a cursor needs the platform, which a view's type does not.
    private static Cursor? s_columnCursor;
    private static Cursor? s_rowCursor;
    private static Cursor? s_fillCursor;

    private Drag _drag;

    // Wheel turns towards a zoom step: a touchpad sends many small ones.
    private double _zoomTurn;

    // The column or row being resized, its size and where the mouse took it, and whether it has been dragged.
    private int _resized;
    private double _resizeSize;
    private double _resizeFrom;
    private bool _resizeMoved;

    // The cells a fill starts from, and those it would fill now.
    private CellArea _fillFrom;
    private CellArea _fillTo;

    /// <inheritdoc />
    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        var properties = e.GetCurrentPoint(this).Properties;
        if (IsInEditor(e.Source))
            return;
        if (properties.IsRightButtonPressed)
        {
            if (IsEditing && !CommitEdit())
                return;
            Focus();
            OpenMenu(e.GetPosition(this));
            e.Handled = true;
            return;
        }

        if (!properties.IsLeftButtonPressed)
            return;

        // While a formula is typed, a click where it takes an operand picks a reference; a click
        // elsewhere enters what is being typed first, and a formula that does not parse stays open.
        if (IsEditing && StartPointing(e.GetPosition(this), e.KeyModifiers.HasFlag(KeyModifiers.Shift)))
        {
            _drag = Drag.Point;
            e.Pointer.Capture(this);
            e.Handled = true;
            return;
        }

        if (IsEditing && !CommitEdit())
            return;
        Focus();
        e.Handled = true;
        var point = e.GetPosition(this);
        if (StartChange(point, e.ClickCount))
        {
            e.Pointer.Capture(this);
            return;
        }

        if (Painter.Hit(Viewport, point.X, point.Y, Worksheet) is not { } hit)
            return;
        if (e.ClickCount == 2 && hit.Part == GridPart.Cell)
        {
            BeginEdit();
            return;
        }

        Selection = e.KeyModifiers.HasFlag(KeyModifiers.Shift) && Selection is { } selection
            ? selection.ExtendTo(hit, Worksheet)
            : GridSelection.At(hit, Worksheet);
        _drag = Drag.Select;
        e.Pointer.Capture(this);
    }

    // A press on the edge of a header or on the fill handle starts changing the sheet: true then.
    private bool StartChange(Point point, int clicks)
    {
        if (IsReadOnly || Worksheet is null)
            return false;
        var viewport = Viewport;
        if (Painter.ColumnEdgeAt(viewport, point.X, point.Y) is { } column)
        {
            if (clicks == 2)
                FitColumns(column);
            else
                StartResize(Drag.ColumnEdge, column, Painter.Layout.ColumnWidth(column), point.X);
            return true;
        }

        if (Painter.RowEdgeAt(viewport, point.X, point.Y) is { } row)
        {
            if (clicks == 2)
                FitRows(row);
            else
                StartResize(Drag.RowEdge, row, Painter.Layout.RowHeight(row), point.Y);
            return true;
        }

        if (Selection is { } selection && Painter.IsOnFillHandle(viewport, point.X, point.Y, selection, Worksheet))
        {
            _drag = Drag.Fill;
            _fillFrom = _fillTo = selection.Area;
            return true;
        }

        return false;
    }

    private void StartResize(Drag drag, int index, double size, double from)
    {
        _drag = drag;
        _resized = index;
        _resizeSize = size;
        _resizeFrom = from;
        _resizeMoved = false;
    }

    /// <inheritdoc />
    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        var point = e.GetPosition(this);
        switch (_drag)
        {
            case Drag.None:
                Cursor = HoverCursor(point);
                return;
            case Drag.Select when Selection is { } selection:
            {
                // Past an edge the nearest cell is the next one out of view, and scrolling to it follows the mouse.
                var cell = Painter.CellNearest(Viewport, point.X, point.Y);
                Selection = selection.ExtendTo(cell, Worksheet);
                ScrollTo(Painter.ScrollIntoView(Viewport, cell));
                break;
            }

            case Drag.ColumnEdge:
                Painter.Layout.SetColumnWidth(_resized, Math.Round(Math.Max(0, _resizeSize + point.X - _resizeFrom)));
                _resizeMoved = true;
                InvalidateVisual();
                break;
            case Drag.RowEdge:
                Painter.Layout.SetRowHeight(_resized, Math.Round(Math.Max(0, _resizeSize + point.Y - _resizeFrom)));
                _resizeMoved = true;
                InvalidateVisual();
                break;
            case Drag.Point:
            {
                var cell = Painter.CellNearest(Viewport, point.X, point.Y);
                PointTo(cell);
                ScrollTo(Painter.ScrollIntoView(Viewport, cell));
                break;
            }

            case Drag.Fill:
            {
                var cell = Painter.CellNearest(Viewport, point.X, point.Y);
                _fillTo = GridEditor.FillArea(_fillFrom, cell);
                Painter.FillArea = _fillTo == _fillFrom ? null : _fillTo;
                ScrollTo(Painter.ScrollIntoView(Viewport, cell));
                InvalidateVisual();
                break;
            }
        }

        e.Handled = true;
    }

    /// <inheritdoc />
    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (_drag == Drag.None)
            return;
        var drag = _drag;
        _drag = Drag.None;
        e.Pointer.Capture(null);
        e.Handled = true;
        if (drag == Drag.Point)
        {
            _box?.Focus();
            return;
        }

        if (Worksheet is not { } sheet || Editor is not { } editor)
            return;
        switch (drag)
        {
            case Drag.ColumnEdge when _resizeMoved:
                Resize(sheet, editor, rows: false, Painter.Layout.ColumnWidth(_resized));
                break;
            case Drag.RowEdge when _resizeMoved:
                Resize(sheet, editor, rows: true, Painter.Layout.RowHeight(_resized));
                break;
            case Drag.Fill:
                Painter.FillArea = null;
                InvalidateVisual();
                if (_fillTo != _fillFrom && Try(() => editor.Fill(sheet, _fillFrom, _fillTo)) is not null)
                    Go(GridSelection.Cell(new CellPosition(_fillTo.FirstRow, _fillTo.FirstColumn), sheet).ExtendTo(new CellPosition(_fillTo.LastRow, _fillTo.LastColumn), sheet), extent: true);
                break;
        }
    }

    /// <inheritdoc />
    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e);
        if (_drag is Drag.ColumnEdge or Drag.RowEdge)
            ReloadLayout();
        if (_drag == Drag.Fill)
        {
            Painter.FillArea = null;
            InvalidateVisual();
        }

        _drag = Drag.None;
    }

    /// <inheritdoc />
    protected override void OnPointerWheelChanged(PointerWheelEventArgs e)
    {
        base.OnPointerWheelChanged(e);
        e.Handled = true;

        // Ctrl or Cmd with the wheel zooms in steps of 10% a turn, as in Excel.
        if (CommandKeys.IsCommand(e.KeyModifiers))
        {
            _zoomTurn += e.Delta.Y;
            var steps = (int)Math.Truncate(_zoomTurn);
            if (steps != 0)
            {
                _zoomTurn -= steps;
                Zoom = Math.Round(Zoom * 10 + steps) / 10;
            }

            return;
        }

        var layout = Painter.Layout;
        var step = new Vector(e.Delta.X * WheelRows * layout.DefaultColumnWidth, e.Delta.Y * WheelRows * layout.DefaultRowHeight);
        ScrollTo(Viewport with { OffsetX = Viewport.OffsetX - step.X, OffsetY = Viewport.OffsetY - step.Y });
    }

    private Cursor? HoverCursor(Point point)
    {
        if (IsReadOnly || Worksheet is null)
            return null;
        var viewport = Viewport;
        if (Painter.ColumnEdgeAt(viewport, point.X, point.Y) is not null)
            return s_columnCursor ??= new Cursor(StandardCursorType.SizeWestEast);
        if (Painter.RowEdgeAt(viewport, point.X, point.Y) is not null)
            return s_rowCursor ??= new Cursor(StandardCursorType.SizeNorthSouth);
        return Painter.IsOnFillHandle(viewport, point.X, point.Y, Selection, Worksheet) ? s_fillCursor ??= new Cursor(StandardCursorType.Cross) : null;
    }

    // A column or row dragged to a size: with whole columns or rows selected and it among them,
    // all of them take it, as in Excel. Dragged to nothing, they are hidden.
    private void Resize(Worksheet sheet, GridEditor editor, bool rows, double pixels)
    {
        var (first, last) = Resized(rows, _resized);
        if (pixels < 1)
            Try(() => rows ? editor.SetRowHidden(sheet, first, last, true) : editor.SetColumnHidden(sheet, first, last, true));
        else if (rows)
            Try(() => editor.SetRowHeight(sheet, first, last, SheetLayout.RowHeightFromPixels(pixels, Painter.Zoom)));
        else
            Try(() => editor.SetColumnWidth(sheet, first, last, SheetLayout.ColumnWidthFromPixels(pixels, Painter.MaxDigitWidth(sheet))));
        ReloadLayout();
    }

    private (int First, int Last) Resized(bool rows, int index)
    {
        if (Selection is { } selection && (rows ? selection.IsEntireRows : selection.IsEntireColumns))
        {
            var area = selection.Area;
            var (first, last) = rows ? (area.FirstRow, area.LastRow) : (area.FirstColumn, area.LastColumn);
            if (index >= first && index <= last)
                return (first, last);
        }

        return (index, index);
    }

    private void FitColumns(int column)
    {
        if (Worksheet is not { } sheet || Editor is not { } editor)
            return;
        var (first, last) = Resized(rows: false, column);
        Try(() => editor.FitColumnWidth(sheet, first, last, Painter));
    }

    private void FitRows(int row)
    {
        if (Worksheet is not { } sheet || Editor is not { } editor)
            return;
        var (first, last) = Resized(rows: true, row);
        Try(() => editor.FitRowHeight(sheet, first, last, Painter));
    }
}
