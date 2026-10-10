using System.Globalization;

namespace SharpCell.Grid.Tests.Selection;

public class GridEditorTests
{
    private static (Worksheet Sheet, GridEditor Editor) Sheet()
    {
        var workbook = new Workbook { Culture = CultureInfo.GetCultureInfo("en-US") };
        var sheet = workbook.AddSheet("S");
        sheet["A1"].Value = 2;
        sheet["B1"].Formula = "=A1*10";
        workbook.Recalculate();
        return (sheet, new GridEditor(workbook));
    }

    private static CellArea Area(string first, string? last = null)
    {
        var (a, b) = (Cells.Parse(first), Cells.Parse(last ?? first));
        return new CellArea(a.Row, a.Column, b.Row, b.Column);
    }

    [Fact]
    public void Typing_reads_the_text_as_Excel_and_recalculates()
    {
        var (sheet, editor) = Sheet();

        var edit = editor.Enter(sheet, Area("A1"), "5");

        Assert.Equal(50.0, sheet["B1"].Value.AsNumber());
        Assert.Contains(edit.Recalculation.ChangedCells, c => c.Address == "B1");
        Assert.Equal(Area("A1"), edit.Area);
    }

    [Fact]
    public void Typing_into_a_selection_fills_every_cell_but_the_inner_cells_of_merges()
    {
        var (sheet, editor) = Sheet();
        sheet.Merge("C2:D2");

        editor.Enter(sheet, Area("C1", "D2"), "7%");

        Assert.Equal(0.07, sheet["C1"].Value.AsNumber());
        Assert.Equal(0.07, sheet["D1"].Value.AsNumber());
        Assert.Equal(0.07, sheet["C2"].Value.AsNumber());
        Assert.Equal(CellValue.Empty, sheet["D2"].Value);
        Assert.Equal("0%", sheet["C1"].NumberFormat);
    }

    [Fact]
    public void Undo_gives_back_values_formulas_and_formats_and_redo_takes_them_again()
    {
        var (sheet, editor) = Sheet();
        editor.Enter(sheet, Area("A1", "B1"), "12%");
        Assert.Equal("0%", sheet["B1"].NumberFormat);

        var undone = editor.Undo();

        Assert.NotNull(undone);
        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
        Assert.Equal("=A1*10", sheet["B1"].Formula);
        Assert.Equal(20.0, sheet["B1"].Value.AsNumber());
        Assert.Equal("General", sheet["B1"].NumberFormat);
        Assert.True(editor.CanRedo);

        editor.Redo();

        Assert.Equal(0.12, sheet["B1"].Value.AsNumber());
        Assert.Null(sheet["B1"].Formula);
    }

    [Fact]
    public void A_new_change_drops_what_could_be_redone()
    {
        var (sheet, editor) = Sheet();
        editor.Enter(sheet, Area("A1"), "3");
        editor.Undo();

        editor.Enter(sheet, Area("A1"), "4");

        Assert.False(editor.CanRedo);
        Assert.Null(editor.Redo());
    }

    [Fact]
    public void A_broken_formula_changes_no_cell()
    {
        var (sheet, editor) = Sheet();

        Assert.Throws<FormulaParseException>(() => editor.Enter(sheet, Area("A1", "A3"), "=1+"));

        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
        Assert.False(editor.CanUndo);
    }

    [Fact]
    public void Delete_clears_contents_and_keeps_formats()
    {
        var (sheet, editor) = Sheet();
        sheet["A1"].NumberFormat = "0.00";

        editor.Clear(sheet, new CellArea(1, 1, SheetLayout.MaxRow, 1));

        Assert.Equal(CellValue.Empty, sheet["A1"].Value);
        Assert.Equal("0.00", sheet["A1"].NumberFormat);
        Assert.Equal(0.0, sheet["B1"].Value.AsNumber());

        editor.Undo();
        Assert.Equal(2.0, sheet["A1"].Value.AsNumber());
    }

    [Fact]
    public void Pasting_text_from_Excel_reads_each_cell_as_typed()
    {
        var (sheet, editor) = Sheet();

        var edit = editor.Paste(sheet, Cells.Parse("C3"), "Item\t1,250.50\r\n\"Two\nlines\"\t=C4&\"!\"\r\n");

        Assert.Equal(Area("C3", "D4"), edit.Area);
        Assert.Equal(CellValue.Text("Item"), sheet["C3"].Value);
        Assert.Equal(1250.5, sheet["D3"].Value.AsNumber());
        Assert.Equal("#,##0.00", sheet["D3"].NumberFormat);
        Assert.Equal(CellValue.Text("Two\nlines"), sheet["C4"].Value);
        Assert.Equal("Two\nlines!", sheet["D4"].Value.AsText());
    }

    [Fact]
    public void A_pasted_formula_that_does_not_parse_stays_text()
    {
        var (sheet, editor) = Sheet();

        editor.Paste(sheet, Cells.Parse("C1"), "=1+\r\n");

        Assert.Equal(CellValue.Text("=1+"), sheet["C1"].Value);
    }

    [Fact]
    public void Undoing_a_change_over_a_spill_lets_it_spill_again()
    {
        var (sheet, editor) = Sheet();
        sheet["E1"].Formula = "=SEQUENCE(3)";
        sheet.Workbook.Recalculate();

        editor.Enter(sheet, Area("E2"), "x");
        Assert.Equal(ErrorKind.Spill, sheet["E1"].Value.AsError());

        editor.Undo();

        Assert.Equal(2.0, sheet["E2"].Value.AsNumber());
        Assert.Equal(1.0, sheet["E1"].Value.AsNumber());
    }

    [Fact]
    public void Every_change_is_announced()
    {
        var (sheet, editor) = Sheet();
        var edits = new List<GridEdit>();
        editor.Edited += (_, e) => edits.Add(e);

        editor.Enter(sheet, Area("A1"), "1");
        editor.Undo();
        editor.Redo();

        Assert.Equal(3, edits.Count);
    }

    [Theory]
    [InlineData("a\tb\r\nc\td\r\n", "a|b;c|d")]
    [InlineData("a\tb\nc\td", "a|b;c|d")]
    [InlineData("\"x\ty\"\t\"say \"\"hi\"\"\"\r\n", "x\ty|say \"hi\"")]
    [InlineData("\"not closed\tz\r\n", "\"not closed|z")]
    [InlineData("\t\r\n", "|")]
    [InlineData("", "")]
    public void Clipboard_text_reads_as_Excel_reads_it(string text, string expected)
    {
        var rows = GridClipboard.ReadText(text);

        Assert.Equal(expected, string.Join(";", rows.Select(r => string.Join("|", r))));
    }

    [Fact]
    public void Copied_text_reads_back_as_the_same_cells()
    {
        var sheet = new Workbook().AddSheet("S");
        sheet["A1"].Value = "Two\nlines";
        sheet["B1"].Value = "tab\there";
        sheet["A2"].Value = "say \"hi\"";

        var rows = GridClipboard.ReadText(GridClipboard.Text(sheet, Area("A1", "B2")));

        Assert.Equal(["Two\nlines", "tab\there"], rows[0]);
        Assert.Equal(["say \"hi\"", ""], rows[1]);
    }
}
