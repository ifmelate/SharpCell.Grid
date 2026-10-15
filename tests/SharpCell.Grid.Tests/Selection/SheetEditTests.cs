using System.Globalization;
using SharpCell.Xlsx;

namespace SharpCell.Grid.Tests.Selection;

public class SheetEditTests
{
    private static (Worksheet Sheet, GridEditor Editor) Budget()
    {
        var workbook = new Workbook { Culture = CultureInfo.GetCultureInfo("en-US") };
        var sheet = workbook.AddSheet("S");
        var other = workbook.AddSheet("Other");
        sheet["A1"].Input = "Item";
        sheet["A1"].Style = CellStyle.Default with { Font = CellFont.Default with { Bold = true } };
        for (var row = 2; row <= 5; row++)
        {
            sheet[row, 1].Input = "Line " + row;
            sheet[row, 2].Input = (row * 10).ToString(CultureInfo.InvariantCulture);
        }

        sheet["B6"].Input = "=SUM(B2:B5)";
        sheet["C3"].Input = "=B3*2";
        other["A1"].Input = "=S!B4+S!B3";
        workbook.DefineName("Third", "=S!$B$3");
        sheet.Merge("D3:E3");
        sheet.SetRowHeight(3, 30);
        sheet["B3"].NumberFormat = "0.00";
        workbook.Recalculate();
        return (sheet, new GridEditor(workbook));
    }

    private static CellArea Area(string first, string? last = null)
    {
        var (a, b) = (Cells.Parse(first), Cells.Parse(last ?? first));
        return new CellArea(a.Row, a.Column, b.Row, b.Column);
    }

    [Fact]
    public void Inserted_rows_move_cells_down_take_the_formats_above_and_undo_takes_them_out()
    {
        var (sheet, editor) = Budget();
        sheet["B2"].NumberFormat = "0%";

        var edit = editor.InsertRows(sheet, 3, 2);

        Assert.True(edit.ChangesLayout);
        Assert.Equal(new CellArea(3, 1, 4, SheetLayout.MaxColumn), edit.Area);
        Assert.Equal("=SUM(B2:B7)", sheet["B8"].Formula);
        Assert.Equal(("0%", "0%"), (sheet["B3"].NumberFormat, sheet["B4"].NumberFormat));
        Assert.Equal("0.00", sheet["B5"].NumberFormat);
        Assert.Equal(CellValue.Empty, sheet["A3"].Value);

        editor.Undo();
        Assert.Equal("=SUM(B2:B5)", sheet["B6"].Formula);
        Assert.Equal("Line 3", sheet["A3"].Text);

        editor.Redo();
        Assert.Equal("Line 3", sheet["A5"].Text);
    }

    [Fact]
    public void Undoing_a_delete_gives_back_the_cells_sizes_formulas_names_and_merges()
    {
        var (sheet, editor) = Budget();
        var other = sheet.Workbook["Other"];

        editor.DeleteRows(sheet, 3, 1);

        Assert.Equal("=SUM(B2:B4)", sheet["B5"].Formula);
        Assert.Equal("=S!B3+S!#REF!", other["A1"].Formula);
        Assert.Empty(sheet.MergedAreas);

        editor.Undo();

        Assert.Equal("Line 3", sheet["A3"].Text);
        Assert.Equal("30.00", sheet["B3"].Text);
        Assert.Equal("=B3*2", sheet["C3"].Formula);
        Assert.Equal("=SUM(B2:B5)", sheet["B6"].Formula);
        Assert.Equal("=S!B4+S!B3", other["A1"].Formula);
        Assert.Equal("=S!$B$3", sheet.Workbook.DefinedNames.Single(n => n.Name == "Third").Formula);
        Assert.Equal(30, sheet.RowHeight(3));
        Assert.Equal(["D3:E3"], sheet.MergedAreas.Select(m => m.Address));
        Assert.Equal(140, sheet["B6"].Value.AsNumber());
        Assert.Equal(70, other["A1"].Value.AsNumber());

        editor.Redo();
        Assert.Equal("Line 4", sheet["A3"].Text);
    }

    [Fact]
    public void Columns_are_inserted_and_deleted_alike()
    {
        var (sheet, editor) = Budget();

        editor.InsertColumns(sheet, 2, 1);
        Assert.Equal("=C3*2", sheet["D3"].Formula);
        editor.DeleteColumns(sheet, 1, 1);
        Assert.Equal("=B3*2", sheet["C3"].Formula);
        Assert.Equal(CellValue.Empty, sheet["A3"].Value);

        editor.Undo();
        editor.Undo();
        Assert.Equal("=B3*2", sheet["C3"].Formula);
        Assert.Equal("Line 3", sheet["A3"].Text);
        Assert.True(sheet["A1"].Style.Font.Bold);
    }

    [Fact]
    public void Widths_heights_and_hiding_are_steps_to_undo()
    {
        var (sheet, editor) = Budget();

        editor.SetColumnWidth(sheet, 1, 2, 20);
        editor.SetRowHeight(sheet, 2, 2, 40);
        editor.SetColumnHidden(sheet, 3, 3, true);
        editor.SetRowHidden(sheet, 4, 5, true);

        Assert.Equal((20, 20), (sheet.ColumnWidth(1), sheet.ColumnWidth(2)));
        Assert.Equal(40, sheet.RowHeight(2));
        Assert.True(sheet.IsColumnHidden(3));
        Assert.True(sheet.IsRowHidden(5));

        editor.Undo();
        editor.Undo();
        editor.Undo();
        editor.Undo();
        Assert.Null(sheet.ColumnWidth(1));
        Assert.Null(sheet.RowHeight(2));
        Assert.False(sheet.IsColumnHidden(3));
        Assert.False(sheet.IsRowHidden(5));
    }

    [Fact]
    public void Copying_moves_formula_references_with_the_cells_and_repeats_over_a_larger_area()
    {
        var (sheet, editor) = Budget();

        var edit = editor.Copy(sheet, Area("C3"), Area("C4", "C5"));

        Assert.Equal(Area("C4", "C5"), edit.Area);
        Assert.Equal("=B5*2", sheet["C5"].Formula);
        Assert.Equal(100, sheet["C5"].Value.AsNumber());

        editor.Undo();
        Assert.Null(sheet["C5"].Formula);
    }

    [Fact]
    public void A_copy_cut_by_the_sheet_edge_keeps_what_fits()
    {
        var (sheet, editor) = Budget();

        var edit = editor.Copy(sheet, Area("A2", "A3"), Area("A1048576"));

        Assert.Equal(Area("A1048576"), edit.Area);
        Assert.Equal("Line 2", sheet["A1048576"].Text);
    }

    [Fact]
    public void Sizes_hidden_rows_and_inserted_rows_set_in_the_grid_are_saved()
    {
        var (sheet, editor) = Budget();
        editor.SetColumnWidth(sheet, 1, 1, 20);
        editor.SetRowHidden(sheet, 4, 4, true);
        editor.InsertRows(sheet, 2, 1);

        using var stream = new MemoryStream();
        XlsxWriter.Save(sheet.Workbook, stream);
        stream.Position = 0;
        var saved = XlsxReader.Load(stream)["S"];

        Assert.Equal(20, saved.ColumnWidth(1));
        Assert.True(saved.IsRowHidden(5));
        Assert.Equal("=SUM(B3:B6)", saved["B7"].Formula);
    }

    [Fact]
    public void Undoing_a_delete_gives_back_validation_and_links_of_the_deleted_rows()
    {
        var (sheet, editor) = Budget();
        sheet.AddValidation("B2:B5", DataValidation.Custom("=B2>0"));
        sheet["A3"].Hyperlink = new Hyperlink("#Other!A1");

        editor.DeleteRows(sheet, 2, 2);
        Assert.Equal(["B2:B3"], sheet.Validations.Select(v => v.Range.Address));
        editor.Undo();

        var (range, rule) = Assert.Single(sheet.Validations);
        Assert.Equal("B2:B5", range.Address);
        Assert.Equal("=B2>0", rule.Formula1);
        Assert.Equal("#Other!A1", sheet["A3"].Hyperlink?.Target);
    }
    [Fact]
    public void Cells_inserted_with_a_shift_down_move_their_column_and_take_the_format_above()
    {
        var (sheet, editor) = Budget();
        sheet["B2"].NumberFormat = "0%";

        var edit = editor.InsertCells(sheet, Area("B3"), InsertShift.Down);

        Assert.Equal(Area("B3"), edit.Area);
        Assert.Equal((CellValue.Empty, "0%"), (sheet["B3"].Value, sheet["B3"].NumberFormat));
        Assert.Equal(30.0, sheet["B4"].Value.AsNumber());
        Assert.Equal("Line 3", sheet["A3"].Text);                    // column A stays
        Assert.Equal("=SUM(B2:B6)", sheet["B7"].Formula);
        Assert.Equal("=S!B5+S!B4", sheet.Workbook["Other"]["A1"].Formula);

        editor.Undo();
        Assert.Equal(("=SUM(B2:B5)", 30.0), (sheet["B6"].Formula, sheet["B3"].Value.AsNumber()));
        editor.Redo();
        Assert.Equal(CellValue.Empty, sheet["B3"].Value);
    }

    [Fact]
    public void Cells_deleted_with_a_shift_left_come_back_whole_on_undo()
    {
        var (sheet, editor) = Budget();

        editor.DeleteCells(sheet, Area("B3"), DeleteShift.Left);

        Assert.Equal("=#REF!*2", sheet["B3"].Formula);                  // C3 moved into B3, and B3 is gone
        Assert.Equal("=S!B4+S!#REF!", sheet.Workbook["Other"]["A1"].Formula);
        Assert.Equal(40.0, sheet["B4"].Value.AsNumber());

        editor.Undo();
        Assert.Equal((30.0, "0.00", "=B3*2"), (sheet["B3"].Value.AsNumber(), sheet["B3"].NumberFormat, sheet["C3"].Formula));
        Assert.Equal("=S!B4+S!B3", sheet.Workbook["Other"]["A1"].Formula);
        Assert.Equal(60.0, sheet["C3"].Value.AsNumber());
    }

    [Fact]
    public void Whole_rows_insert_rows_and_a_merged_area_across_the_cells_refuses()
    {
        var (sheet, editor) = Budget();

        var edit = editor.InsertCells(sheet, new CellArea(3, 1, 3, SheetLayout.MaxColumn), InsertShift.Down);
        Assert.Equal(30.0, sheet.RowHeight(4));                          // the row moved with its height
        Assert.True(edit.ChangesLayout);

        Assert.Throws<InvalidOperationException>(() => editor.DeleteCells(sheet, Area("D2"), DeleteShift.Up));   // D4:E4 merged
        Assert.Equal("Line 3", sheet["A4"].Text);
    }
}
