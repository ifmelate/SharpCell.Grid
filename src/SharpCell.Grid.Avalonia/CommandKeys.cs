using Avalonia.Controls;
using Avalonia.Input;

namespace SharpCell.Grid.Avalonia;

/// <summary>
/// Ctrl and Cmd as one: every shortcut works with either, on every platform. In a browser the
/// platform is the browser, whose gestures use Ctrl even on a Mac, where people press Cmd.
/// </summary>
internal static class CommandKeys
{
    // Ctrl or Cmd without Alt: Ctrl+Alt is AltGr on Windows, which types characters.
    public static bool IsCommand(KeyModifiers modifiers) =>
        (modifiers.HasFlag(KeyModifiers.Control) || modifiers.HasFlag(KeyModifiers.Meta)) && !modifiers.HasFlag(KeyModifiers.Alt);

    /// <summary>The editing shortcuts of a text box with Ctrl or Cmd; whether the key was one.</summary>
    public static bool TextBoxKey(TextBox box, KeyEventArgs e)
    {
        if (!IsCommand(e.KeyModifiers))
            return false;
        var shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        switch (e.Key)
        {
            case Key.A:
                box.SelectAll();
                break;
            case Key.C or Key.Insert:
                box.Copy();
                break;
            case Key.X:
                box.Cut();
                break;
            case Key.V:
                box.Paste();
                break;
            case Key.Z when shift:
            case Key.Y:
                box.Redo();
                break;
            case Key.Z:
                box.Undo();
                break;
            default:
                return false;
        }

        return true;
    }
}
