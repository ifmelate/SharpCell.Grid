using System.Runtime.Versioning;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Browser;

[assembly: SupportedOSPlatform("browser")]

namespace SharpCell.Grid.Demo;

internal sealed partial class Program
{
    // The browser has no fonts of its own for the canvas: Inter is bundled, and cells in fonts the
    // demo does not have (Calibri, Arial) are drawn in it.
    private static Task Main(string[] args) => BuildAvaloniaApp().WithInterFont().StartBrowserAppAsync("out");

    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<DemoApp>();
}
