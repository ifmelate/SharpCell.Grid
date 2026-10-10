namespace SharpCell.Grid.Demo;

/// <summary>The workbook the demo opens with: styles, a merged title, frozen panes and formulas SharpCell calculates.</summary>
internal static class Showcase
{
    private static readonly string[] Months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    private static readonly string[] Regions = ["North", "South", "East", "West"];

    public static Workbook Build()
    {
        var workbook = new Workbook();
        Sales(workbook.AddSheet("Sales"));
        About(workbook.AddSheet("About"));
        workbook.Recalculate();
        return workbook;
    }

    private static void Sales(Worksheet sheet)
    {
        var blue = CellColor.FromRgb(0x4472C4);
        var title = CellStyle.Default with
        {
            Font = CellFont.Default with { Size = 16, Bold = true, Color = CellColor.FromRgb(0xFFFFFF) },
            Fill = blue,
            HorizontalAlignment = CellHorizontalAlignment.Center,
            VerticalAlignment = CellVerticalAlignment.Center,
        };
        var header = CellStyle.Default with
        {
            Font = CellFont.Default with { Bold = true },
            Fill = CellColor.FromRgb(0xD9E1F2),
            BottomBorder = new CellBorder(CellBorderStyle.Medium, blue),
            HorizontalAlignment = CellHorizontalAlignment.Center,
            WrapText = true,
        };
        var total = CellStyle.Default with
        {
            Font = CellFont.Default with { Bold = true },
            TopBorder = new CellBorder(CellBorderStyle.Thin),
            BottomBorder = new CellBorder(CellBorderStyle.Double),
        };

        sheet["A1"].Value = "Sales by region, 2026";
        sheet["A1"].Style = title;
        sheet.Merge("A1:H1");
        sheet.SetRowHeight(1, 30);

        string[] titles = ["Month", .. Regions, "Total", "Share of the year", "Note"];
        for (var i = 0; i < titles.Length; i++)
        {
            sheet[3, i + 1].Value = titles[i];
            sheet[3, i + 1].Style = header;
        }

        sheet.SetRowHeight(3, 30);
        sheet.SetColumnWidth(1, 14);
        sheet.SetColumnWidth(7, 12);

        for (var m = 0; m < Months.Length; m++)
        {
            var row = 4 + m;
            sheet[row, 1].Value = Months[m];
            for (var r = 0; r < Regions.Length; r++)
            {
                // Seasonal figures that differ by region, the same on every run.
                sheet[row, 2 + r].Value = 900 + 40 * ((m * 7 + r * 5) % 13) + 150 * r;
                sheet[row, 2 + r].NumberFormat = "#,##0";
            }

            sheet[row, 6].Formula = $"=SUM(B{row}:E{row})";
            sheet[row, 6].NumberFormat = "#,##0";
            sheet[row, 7].Formula = $"=F{row}/$F$16";
            sheet[row, 7].NumberFormat = "0.0%";
            sheet[row, 8].Formula = $"=IF(F{row}=MAX($F$4:$F$15),\"The best month of the year\",IF(F{row}=MIN($F$4:$F$15),\"The weakest month\",\"\"))";
            if (m % 2 == 1)
            {
                for (var c = 1; c <= 7; c++)
                    sheet[row, c].Style = CellStyle.Default with { Fill = CellColor.FromRgb(0xF2F2F2) };
            }
        }

        sheet["A16"].Value = "Total";
        sheet["A16"].Style = total;
        for (var c = 2; c <= 7; c++)
        {
            var column = (char)('A' + c - 1);
            sheet[16, c].Formula = $"=SUM({column}4:{column}15)";
            sheet[16, c].NumberFormat = c == 7 ? "0.0%" : "#,##0";
            sheet[16, c].Style = total;
        }

        sheet["A18"].Value = "Figures are made up. Totals, shares and notes are formulas calculated by SharpCell in your browser.";
        sheet["A18"].Style = CellStyle.Default with { Font = CellFont.Default with { Italic = true, Color = CellColor.FromRgb(0x7F7F7F) } };
        sheet.FrozenRows = 3;
        sheet.FrozenColumns = 1;
    }

    private static void About(Worksheet sheet)
    {
        string[] lines =
        [
            "SharpCell.Grid is a free (MIT) Excel-like spreadsheet control for .NET, built on the SharpCell formula engine.",
            "It shows number formats, cell styles, column widths, row heights, merged cells, frozen panes and text running on over empty cells.",
            "This page runs it in your browser with WebAssembly: an .xlsx file you open is read here and is never uploaded.",
            "Type into a cell, press F2 or double-click to edit, Delete to clear, Ctrl+Z to undo; formulas recalculate as you go.",
            "Select with the mouse and keyboard as in Excel, copy to Excel with Ctrl+C and paste from it with Ctrl+V.",
        ];
        for (var i = 0; i < lines.Length; i++)
            sheet[1 + 2 * i, 1].Value = lines[i];
        sheet["A1"].Style = CellStyle.Default with { Font = CellFont.Default with { Bold = true, Size = 13 } };
    }
}
