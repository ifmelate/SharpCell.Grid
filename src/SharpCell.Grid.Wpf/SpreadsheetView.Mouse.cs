using System;
using System.Windows;
using System.Windows.Input;

namespace SharpCell.Grid.Wpf;

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
    }

    private Drag _drag;

    // Wheel notches towards a zoom step: a touchpad sends many small turns.
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
    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnMouseLeftButtonDown(e);
        if (IsInEditor(e.OriginalSource))
            return;

        // A click elsewhere enters what is being typed first; a formula that does not parse stays open.
        if (IsEditing && !CommitEdit())
            return;
        Focus();
        e.Handled = true;
        if (Press(e.GetPosition(this), Keyboard.Modifiers.HasFlag(ModifierKeys.Shift), e.ClickCount))
            CaptureMouse();
    }

    /// <inheritdoc />
    protected override void OnMouseRightButtonDown(MouseButtonEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnMouseRightButtonDown(e);
        if (IsInEditor(e.OriginalSource) || (IsEditing && !CommitEdit()))
            return;
        Focus();
        if (OpenMenu(e.GetPosition(this)))
            e.Handled = true;
    }

    /// <inheritdoc />
    protected override void OnMouseMove(MouseEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnMouseMove(e);
        if (_drag == Drag.None)
        {
            Cursor = HoverCursor(e.GetPosition(this));
            return;
        }

        if (IsMouseCaptured)
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
        if (_drag == Drag.None)
            return;
        Release();
        ReleaseMouseCapture();
        e.Handled = true;
    }

    /// <inheritdoc />
    protected override void OnLostMouseCapture(MouseEventArgs e)
    {
        base.OnLostMouseCapture(e);
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
    protected override void OnMouseWheel(MouseWheelEventArgs e)
    {
        ArgumentNullException.ThrowIfNull(e);
        base.OnMouseWheel(e);
        e.Handled = true;
        Wheel(e.Delta, Keyboard.Modifiers.HasFlag(ModifierKeys.Control));
    }

    // A turn of the wheel scrolls; with Ctrl it zooms in steps of 10% a notch, as in Excel.
    internal void Wheel(double delta, bool control)
    {
        if (control)
        {
            _zoomTurn += delta / WheelNotch;
            var steps = (int)Math.Truncate(_zoomTurn);
            if (steps != 0)
            {
                _zoomTurn -= steps;
                Zoom = Math.Round(Zoom * 10 + steps) / 10;
            }

            return;
        }

        var rows = delta / WheelNotch * WheelRows;
        ScrollTo(Viewport with { OffsetY = Viewport.OffsetY - rows * Painter.Layout.DefaultRowHeight });
    }

    // What a press of the left button does: resizing on the edge of a header, filling on the fill
    // handle, selecting elsewhere, editing on a double-click of a cell; whether a drag began.
    internal bool Press(Point point, bool shift, int clicks = 1)
    {
        if (!IsReadOnly && Worksheet is not null && StartChange(point, clicks))
            return _drag != Drag.None;
        if (Painter.Hit(Viewport, point.X, point.Y, Worksheet) is not { } hit)
            return false;
        if (clicks == 2 && hit.Part == GridPart.Cell)
        {
            BeginEdit();
            return false;
        }

        Selection = shift && Selection is { } selection ? selection.ExtendTo(hit, Worksheet) : GridSelection.At(hit, Worksheet);
        _drag = Drag.Select;
        return true;
    }

    // Past an edge the nearest cell is the next one out of view, and scrolling to it follows the mouse.
    internal void DragTo(Point point)
    {
        switch (_drag)
        {
            case Drag.Select when Selection is { } selection:
            {
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
    }

    // The button goes up: a resize or a fill is done.
    internal void Release()
    {
        var drag = _drag;
        _drag = Drag.None;
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

    // A press on the edge of a header or on the fill handle starts changing the sheet: true then.
    private bool StartChange(Point point, int clicks)
    {
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

    private Cursor? HoverCursor(Point point)
    {
        if (IsReadOnly || Worksheet is null)
            return null;
        var viewport = Viewport;
        if (Painter.ColumnEdgeAt(viewport, point.X, point.Y) is not null)
            return Cursors.SizeWE;
        if (Painter.RowEdgeAt(viewport, point.X, point.Y) is not null)
            return Cursors.SizeNS;
        return Painter.IsOnFillHandle(viewport, point.X, point.Y, Selection, Worksheet) ? Cursors.Cross : null;
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
