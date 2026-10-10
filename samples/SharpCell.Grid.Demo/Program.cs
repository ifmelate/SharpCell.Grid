using System.Runtime.Versioning;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Browser;
using Avalonia.Media;

[assembly: SupportedOSPlatform("browser")]

namespace SharpCell.Grid.Demo;

internal sealed partial class Program
{
    // The browser has no fonts of its own for the canvas: Inter is bundled and made the default, so
    // cells in fonts the demo does not have (Calibri, Arial) are drawn in it rather than in a
    // monospaced fallback.
    private static Task Main(string[] args) => BuildAvaloniaApp()
        .WithInterFont()
        .With(new FontManagerOptions { DefaultFamilyName = "fonts:Inter#Inter" })
        .StartBrowserAppAsync("out");

    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<DemoApp>();
}
