using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Layout;

namespace SharpCell.Grid.Avalonia;

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
        _cellsDialog = new CellsDialog(insert, wide ? 1 : 0, choice =>
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
                    Structure(a => insert ? Editor!.InsertRows(Worksheet!, a.FirstRow, a.LastRow - a.FirstRow + 1) : Editor!.DeleteRows(Worksheet!, a.FirstRow, a.LastRow - a.FirstRow + 1));
                    break;
                case 3:
                    Structure(a => insert ? Editor!.InsertColumns(Worksheet!, a.FirstColumn, a.LastColumn - a.FirstColumn + 1) : Editor!.DeleteColumns(Worksheet!, a.FirstColumn, a.LastColumn - a.FirstColumn + 1));
                    break;
            }
        }, () =>
        {
            _cellsDialog = null;
            Focus();
        });
        _cellsDialog.Show(this);
    }

    /// <summary>Excel's Insert or Delete dialog: four choices, OK and Cancel, in a flyout over the view.</summary>
    internal sealed class CellsDialog
    {
        private readonly Flyout _flyout;
        private readonly RadioButton[] _choices;
        private readonly Action<int> _ok;
        private readonly Action _cancel;
        private bool _done;

        public CellsDialog(bool insert, int initial, Action<int> ok, Action cancel)
        {
            (_ok, _cancel) = (ok, cancel);
            string[] labels = insert
                ? ["Shift cells right", "Shift cells down", "Entire row", "Entire column"]
                : ["Shift cells left", "Shift cells up", "Entire row", "Entire column"];
            var group = insert ? "insert-cells" : "delete-cells";
            _choices = [.. labels.Select(label => new RadioButton { Content = label, GroupName = group })];
            _choices[initial].IsChecked = true;
            var okButton = new Button { Content = "OK", MinWidth = 72 };
            var cancelButton = new Button { Content = "Cancel", MinWidth = 72 };
            okButton.Click += (_, _) => Choose(Choice);
            cancelButton.Click += (_, _) => Cancel();
            var panel = new StackPanel { Spacing = 4, Margin = new Thickness(8) };
            panel.Children.Add(new TextBlock { Text = insert ? "Insert" : "Delete", FontWeight = global::Avalonia.Media.FontWeight.Bold, Margin = new Thickness(0, 0, 0, 4) });
            foreach (var choice in _choices)
                panel.Children.Add(choice);
            panel.Children.Add(new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Spacing = 8,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 8, 0, 0),
                Children = { okButton, cancelButton },
            });

            // Enter takes the choice and Escape leaves, as in Excel's dialog.
            panel.KeyDown += (_, e) =>
            {
                if (e.Key == Key.Enter)
                    Choose(Choice);
                else if (e.Key == Key.Escape)
                    Cancel();
                else
                    return;
                e.Handled = true;
            };
            _flyout = new Flyout { Content = panel, Placement = PlacementMode.Center };
            _flyout.Closed += (_, _) => Cancel();
        }

        /// <summary>The checked choice: 0 and 1 shift cells, 2 is entire rows, 3 entire columns.</summary>
        public int Choice => Array.FindIndex(_choices, c => c.IsChecked == true);

        public void Show(Control target)
        {
            _flyout.ShowAt(target);
            _choices[Math.Max(0, Choice)].Focus();
        }

        /// <summary>Checks a choice and takes it, as clicking it and OK does.</summary>
        public void Choose(int choice)
        {
            if (_done)
                return;
            _done = true;
            _flyout.Hide();
            _ok(choice);
        }

        public void Cancel()
        {
            if (_done)
                return;
            _done = true;
            _flyout.Hide();
            _cancel();
        }
    }
}
