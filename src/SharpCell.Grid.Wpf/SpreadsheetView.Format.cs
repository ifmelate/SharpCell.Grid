using System;
using System.Windows.Input;

namespace SharpCell.Grid.Wpf;

// Changing how the selected cells look, as Excel's Home tab and its shortcuts do.
public partial class SpreadsheetView
{
    /// <summary>Makes the selected cells bold, or not bold when the active cell is, as Ctrl+B does.</summary>
    public void ToggleBold() => ToggleFont(font => font.Bold, (font, on) => font with { Bold = on });

    /// <summary>Makes the selected cells italic, or not italic when the active cell is, as Ctrl+I does.</summary>
    public void ToggleItalic() => ToggleFont(font => font.Italic, (font, on) => font with { Italic = on });

    /// <summary>Underlines the selected cells, or takes the underline off when the active cell has one, as Ctrl+U does.</summary>
    public void ToggleUnderline() => ToggleFont(font => font.Underline != CellUnderline.None, (font, on) => font with { Underline = on ? CellUnderline.Single : CellUnderline.None });

    /// <summary>Strikes the selected cells through, or not when the active cell is, as Ctrl+5 does.</summary>
    public void ToggleStrikethrough() => ToggleFont(font => font.Strikethrough, (font, on) => font with { Strikethrough = on });

    /// <summary>Fills the selected cells with a colour, or takes their fill off with null, as Excel's Fill Color does.</summary>
    public void SetFill(CellColor? color) => Restyle(style => style with { Fill = color });

    /// <summary>Gives the selected cells' text a colour, or the automatic colour with null, as Excel's Font Color does.</summary>
    public void SetFontColor(CellColor? color) => Restyle(style => style with { Font = style.Font with { Color = color } });

    /// <summary>Puts a border on edges of the selection, or takes it off with null, as Excel's Borders menu does; see <see cref="GridEditor.SetBorders"/>.</summary>
    public void SetBorders(BorderEdges edges, CellBorder? border) => Structure(area => Editor!.SetBorders(Worksheet!, area, edges, border));

    /// <summary>Gives the selected cells a number format, as Excel's Number Format box does.</summary>
    public void SetNumberFormat(string format) => Structure(area => Editor!.SetNumberFormat(Worksheet!, area, format));

    /// <summary>Changes the style of the selected cells: the function gets each cell's style and gives the new one.</summary>
    public void Restyle(Func<CellStyle, CellStyle> change) => Structure(area => Editor!.SetStyle(Worksheet!, area, change));

    // On when the active cell is off, off when it is on: Excel's toggles follow the active cell.
    private void ToggleFont(Func<CellFont, bool> isOn, Func<CellFont, bool, CellFont> set)
    {
        if (Worksheet is not { } sheet || Selection is not { } selection)
            return;
        var on = !isOn(sheet[selection.ActiveCell.Row, selection.ActiveCell.Column].Style.Font);
        Restyle(style => style with { Font = set(style.Font, on) });
    }

    // Excel's formatting shortcuts with Ctrl (or Cmd): B, I, U and 2 to 5 for the font; with Shift,
    // & for an outline border, _ to take borders off, and ~ ! $ % ^ # @ for number formats.
    private bool FormatKey(Key key, bool shift)
    {
        if (!shift)
        {
            switch (key)
            {
                case Key.B or Key.D2:
                    ToggleBold();
                    return true;
                case Key.I or Key.D3:
                    ToggleItalic();
                    return true;
                case Key.U or Key.D4:
                    ToggleUnderline();
                    return true;
                case Key.D5:
                    ToggleStrikethrough();
                    return true;
                default:
                    return false;
            }
        }

        if (key == Key.D7)
        {
            SetBorders(BorderEdges.Outside, new CellBorder(CellBorderStyle.Thin));
            return true;
        }

        if (key == Key.OemMinus)
        {
            SetBorders(BorderEdges.All, null);
            return true;
        }

        var symbol = key switch
        {
            Key.OemTilde => '~',
            Key.D1 => '!',
            Key.D4 => '$',
            Key.D5 => '%',
            Key.D6 => '^',
            Key.D3 => '#',
            Key.D2 => '@',
            _ => '\0',
        };
        if (symbol == '\0' || Worksheet is not { } sheet || GridEditor.ShortcutFormat(symbol, sheet.Workbook.Culture) is not { } format)
            return false;
        SetNumberFormat(format);
        return true;
    }
}
