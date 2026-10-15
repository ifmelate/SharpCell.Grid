using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace SharpCell.Grid.Wpf;

// Inserting and deleting cells with Excel's dialog: shift the cells beside them, or whole rows or columns.
public partial class SpreadsheetView
{
    private CellsDialog? _cellsDialog;

    /// <summary>Inserts empty cells in place of the selection, moving the cells beside it down or right, as Excel's Insert dialog does.</summary>
    /// <param name="shift">Where the cells beside the selection go.</param>
    public void InsertCells(InsertShift shift) => Structure(area => Editor!.InsertCells(Worksheet!, area, shift));

    /// <summary>Deletes the selected cells, moving the cells below or on the right into their place, as Excel's Delete dialog does.</summary>
    /// <param name="shift">Which cells take the place of the deleted ones.</param>
    public void DeleteCells(DeleteShift shift) => Structure(area => Editor!.DeleteCells(Worksheet!, area, shift));

    /// <summary>
    /// Opens Excel's Insert dialog for the selection: shift cells right or down, or insert entire
    /// rows or columns. Ctrl+Plus opens it when the selection is not whole rows or columns.
    /// </summary>
    public void OpenInsertDialog() => OpenCellsDialog(insert: true);

    /// <summary>Opens Excel's Delete dialog for the selection: shift cells left or up, or delete entire rows or columns.</summary>
    public void OpenDeleteDialog() => OpenCellsDialog(insert: false);

    /// <summary>The dialog open now, for tests; null when none is.</summary>
    internal CellsDialog? OpenDialog => _cellsDialog;

    private void OpenCellsDialog(bool insert)
    {
        if (IsReadOnly || Worksheet is null || Selection is not { } selection || Editor is null)
            return;
        var area = selection.Area;
        var wide = area.LastColumn - area.FirstColumn > area.LastRow - area.FirstRow;
        _cellsDialog = new CellsDialog(this, insert, wide ? 1 : 0, choice =>
        {
            _cellsDialog = null;
            Focus();
            switch (choice)
            {
                case 0 when insert:
                    InsertCells(InsertShift.Right);
                    break;
                case 1 when insert:
                    InsertCells(InsertShift.Down);
                    break;
                case 0:
                    DeleteCells(DeleteShift.Left);
                    break;
                case 1:
                    DeleteCells(DeleteShift.Up);
                    break;
                case 2:
                    if (insert)
                        InsertRows();
                    else
                        DeleteRows();
                    break;
                case 3:
                    if (insert)
                        InsertColumns();
                    else
                        DeleteColumns();
                    break;
            }
        }, () =>
        {
            _cellsDialog = null;
            Focus();
        });
    }

    /// <summary>Excel's Insert or Delete dialog: four choices, OK and Cancel, in a popup over the view.</summary>
    internal sealed class CellsDialog
    {
        private readonly Popup _popup;
        private readonly RadioButton[] _choices;
        private readonly Action<int> _ok;
        private readonly Action _cancel;
        private bool _done;

        public CellsDialog(UIElement target, bool insert, int initial, Action<int> ok, Action cancel)
        {
            (_ok, _cancel) = (ok, cancel);
            string[] labels = insert
                ? ["Shift cells right", "Shift cells down", "Entire row", "Entire column"]
                : ["Shift cells left", "Shift cells up", "Entire row", "Entire column"];
            var group = insert ? "insert-cells" : "delete-cells";
            _choices = [.. labels.Select(label => new RadioButton { Content = label, GroupName = group, Margin = new Thickness(0, 2, 0, 2) })];
            _choices[initial].IsChecked = true;
            var okButton = new Button { Content = "OK", MinWidth = 72, IsDefault = true };
            var cancelButton = new Button { Content = "Cancel", MinWidth = 72, IsCancel = true, Margin = new Thickness(8, 0, 0, 0) };
            okButton.Click += (_, _) => Choose(Choice);
            cancelButton.Click += (_, _) => Cancel();
            var panel = new StackPanel { Margin = new Thickness(10) };
            panel.Children.Add(new TextBlock { Text = insert ? "Insert" : "Delete", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 6) });
            foreach (var choice in _choices)
                panel.Children.Add(choice);
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 10, 0, 0) };
            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);
            panel.Children.Add(buttons);

            // Enter takes the choice and Escape leaves, as in Excel's dialog.
            panel.PreviewKeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter)
                    Choose(Choice);
                else if (e.Key == Key.Escape)
                    Cancel();
                else
                    return;
                e.Handled = true;
            };
            var border = new Border { Child = panel, Background = SystemColors.WindowBrush, BorderBrush = SystemColors.ActiveBorderBrush, BorderThickness = new Thickness(1) };
            _popup = new Popup { Child = border, PlacementTarget = target, Placement = PlacementMode.Center, StaysOpen = false, IsOpen = true };
            _popup.Closed += (_, _) => Cancel();
            _choices[initial].Focus();
        }

        /// <summary>The checked choice: 0 and 1 shift cells, 2 is entire rows, 3 entire columns.</summary>
        public int Choice => Array.FindIndex(_choices, c => c.IsChecked == true);

        /// <summary>Takes a choice, as checking it and clicking OK does.</summary>
        public void Choose(int choice)
        {
            if (_done)
                return;
            _done = true;
            _popup.IsOpen = false;
            _ok(choice);
        }

        public void Cancel()
        {
            if (_done)
                return;
            _done = true;
            _popup.IsOpen = false;
            _cancel();
        }
    }
}
