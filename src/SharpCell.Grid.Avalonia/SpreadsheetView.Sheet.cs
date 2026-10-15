using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;

namespace SharpCell.Grid.Avalonia;

// Commands on whole rows and columns, filling, zoom and the context menu, as in Excel.
public partial class SpreadsheetView
{
    /// <summary>How large the cells are shown, 1 for Excel's 100%; from 0.1 to 4.</summary>
    public static readonly StyledProperty<double> ZoomProperty =
        AvaloniaProperty.Register<SpreadsheetView, double>(nameof(Zoom), 1, coerce: (_, zoom) => double.IsNaN(zoom) ? 1 : Math.Clamp(zoom, GridPainter.MinZoom, GridPainter.MaxZoom));

    /// <summary>
    /// How large the cells are shown, 1 for Excel's 100%, from 0.1 to 4; Ctrl or Cmd with the wheel
    /// changes it in steps of 10%. The cell at the top left stays there.
    /// </summary>
    public double Zoom
    {
        get => GetValue(ZoomProperty);
        set => SetValue(ZoomProperty, value);
    }

    /// <summary>Inserts as many rows as are selected before the selection's first row, as Ctrl+Plus does with rows selected.</summary>
    public void InsertRows() => Structure(area => Editor!.InsertRows(Worksheet!, area.FirstRow, area.LastRow - area.FirstRow + 1));

    /// <summary>Inserts as many columns as are selected before the selection's first column.</summary>
    public void InsertColumns() => Structure(area => Editor!.InsertColumns(Worksheet!, area.FirstColumn, area.LastColumn - area.FirstColumn + 1));

    /// <summary>Deletes the selection's rows, as Ctrl+Minus does with rows selected.</summary>
    public void DeleteRows() => Structure(area => Editor!.DeleteRows(Worksheet!, area.FirstRow, area.LastRow - area.FirstRow + 1));

    /// <summary>Deletes the selection's columns.</summary>
    public void DeleteColumns() => Structure(area => Editor!.DeleteColumns(Worksheet!, area.FirstColumn, area.LastColumn - area.FirstColumn + 1));

    /// <summary>Hides the selection's rows, as Ctrl+9 does.</summary>
    public void HideRows() => Structure(area => Editor!.SetRowHidden(Worksheet!, area.FirstRow, area.LastRow, true));

    /// <summary>Shows the hidden rows among the selection's, as Ctrl+Shift+9 does.</summary>
    public void UnhideRows() => Structure(area => Editor!.SetRowHidden(Worksheet!, area.FirstRow, area.LastRow, false));

    /// <summary>Hides the selection's columns, as Ctrl+0 does.</summary>
    public void HideColumns() => Structure(area => Editor!.SetColumnHidden(Worksheet!, area.FirstColumn, area.LastColumn, true));

    /// <summary>Shows the hidden columns among the selection's, as Ctrl+Shift+0 does.</summary>
    public void UnhideColumns() => Structure(area => Editor!.SetColumnHidden(Worksheet!, area.FirstColumn, area.LastColumn, false));

    /// <summary>
    /// Copies the selection's top row into the rows below it, formulas following, as Ctrl+D does;
    /// with one row selected, the row above it into it. Whole columns are filled down to the last used row.
    /// </summary>
    public void FillDown() => FillFrom(rows: true);

    /// <summary>Copies the selection's left column into the columns right of it, as Ctrl+R does; see <see cref="FillDown"/>.</summary>
    public void FillRight() => FillFrom(rows: false);

    // An edit of rows or columns keeps the selection where it is, as in Excel.
    private void Structure(Func<CellArea, GridEdit> change)
    {
        if (IsReadOnly || Worksheet is null || Selection is not { } selection || Editor is null)
            return;
        Try(() => change(selection.Area));
    }

    private void FillFrom(bool rows)
    {
        if (IsReadOnly || Worksheet is not { } sheet || Selection is not { } selection || Editor is not { } editor)
            return;
        Try(() => (rows ? editor.FillDown(sheet, selection.Area) : editor.FillRight(sheet, selection.Area)));
    }

    // The keys of whole rows and columns and of filling, when cells can change.
    private bool SheetKey(KeyEventArgs e, bool command)
    {
        if (!command)
            return false;
        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        var columns = Selection?.IsEntireColumns == true;
        var rows = Selection?.IsEntireRows == true;
        switch (e.Key)
        {
            case Key.D:
                FillDown();
                break;
            case Key.R:
                FillRight();
                break;
            // With cells rather than whole rows or columns selected, Excel asks how.
            case Key.OemPlus or Key.Add:
                if (columns)
                    InsertColumns();
                else if (rows)
                    InsertRows();
                else
                    OpenInsertDialog();
                break;
            case Key.OemMinus or Key.Subtract:
                if (columns)
                    DeleteColumns();
                else if (rows)
                    DeleteRows();
                else
                    OpenDeleteDialog();
                break;
            case Key.D9:
                if (shift)
                    UnhideRows();
                else
                    HideRows();
                break;
            case Key.D0:
                if (shift)
                    UnhideColumns();
                else
                    HideColumns();
                break;
            default:
                return false;
        }

        return true;
    }

    private void OnZoomChanged()
    {
        // The cell at the top left stays there.
        var first = Painter.VisibleArea(Viewport);
        Painter.Zoom = Zoom;
        ReloadLayout();
        var layout = Painter.Layout;
        var frozenWidth = Painter.FrozenColumns == 0 ? 0 : layout.ColumnLeft(Painter.FrozenColumns) + layout.ColumnWidth(Painter.FrozenColumns);
        var frozenHeight = Painter.FrozenRows == 0 ? 0 : layout.RowTop(Painter.FrozenRows) + layout.RowHeight(Painter.FrozenRows);
        ScrollTo(Viewport with
        {
            OffsetX = first.FirstColumn > Painter.FrozenColumns ? layout.ColumnLeft(first.FirstColumn) - frozenWidth : 0,
            OffsetY = first.FirstRow > Painter.FrozenRows ? layout.RowTop(first.FirstRow) - frozenHeight : 0,
        });
        if (_editing)
        {
            StyleLike(Box(), Worksheet![_editCell.Row, _editCell.Column]);
            InvalidateMeasure();
        }
    }

    // Excel's menu for the cells, the column headers or the row headers under the pointer; a
    // click outside the selection selects there first. An application's own ContextMenu or
    // ContextFlyout takes its place.
    private void OpenMenu(Point point)
    {
        if (ContextMenu is not null || ContextFlyout is not null || Painter.Hit(Viewport, point.X, point.Y, Worksheet) is not { } hit)
            return;
        var sheet = Worksheet;
        var selection = Selection;
        var inside = selection is not null && hit.Part switch
        {
            GridPart.ColumnHeader => selection.IsEntireColumns && hit.Cell.Column >= selection.Area.FirstColumn && hit.Cell.Column <= selection.Area.LastColumn,
            GridPart.RowHeader => selection.IsEntireRows && hit.Cell.Row >= selection.Area.FirstRow && hit.Cell.Row <= selection.Area.LastRow,
            GridPart.Corner => selection.IsEntireColumns && selection.IsEntireRows,
            _ => selection.Contains(hit.Cell.Row, hit.Cell.Column),
        };
        if (!inside)
            Selection = GridSelection.At(hit, sheet);

        var menu = new ContextMenu();
        void Add(string header, Action action, bool enabled = true)
        {
            var item = new MenuItem { Header = header, IsEnabled = enabled };
            item.Click += (_, _) => action();
            menu.Items.Add(item);
        }

        var editable = !IsReadOnly && sheet is not null;
        if (editable)
            Add("Cut", () => _ = CutAsync());
        Add("Copy", () => _ = CopyAsync());
        if (editable)
        {
            Add("Paste", () => _ = PasteAsync());
            menu.Items.Add(new Separator());
            switch (hit.Part)
            {
                case GridPart.ColumnHeader:
                    Add("Insert", InsertColumns);
                    Add("Delete", DeleteColumns);
                    break;
                case GridPart.RowHeader:
                    Add("Insert", InsertRows);
                    Add("Delete", DeleteRows);
                    break;
                case GridPart.Cell:
                    Add("Insert...", OpenInsertDialog);
                    Add("Delete...", OpenDeleteDialog);
                    break;
            }

            Add("Clear Contents", ClearSelection);
            if (hit.Part is GridPart.ColumnHeader or GridPart.RowHeader)
            {
                menu.Items.Add(new Separator());
                var rows = hit.Part == GridPart.RowHeader;
                Add("Hide", rows ? HideRows : HideColumns);
                Add("Unhide", rows ? UnhideRows : UnhideColumns);
            }
        }

        menu.Open(this);
    }
}
