using Avalonia;
using Avalonia.Headless;
using Avalonia.Media;
using Avalonia.Themes.Fluent;
using Avalonia.Threading;

[assembly: AssemblyFixture(typeof(SharpCell.Grid.Tests.Controls.HeadlessSession))]

namespace SharpCell.Grid.Tests.Controls;

/// <summary>
/// The one Avalonia application of the test run, rendering with Skia and the bundled Inter font so
/// frames look the same on every machine. Avalonia code runs on its thread through <see cref="Run"/>.
/// The session is never disposed: <see cref="HeadlessUnitTestSession.Dispose"/> called by xUnit after
/// the last test waits forever and the run hangs; its thread is a background thread and ends with
/// the process.
/// </summary>
public sealed class HeadlessSession
{
    private readonly HeadlessUnitTestSession session = HeadlessUnitTestSession.StartNew(typeof(TestApp));

    // A frame still to render when the test ends would render after the application is gone, without
    // its fonts; it is rendered while the test still runs. A test that fails keeps its own failure.
    public Task Run(Action body) => session.Dispatch(() =>
    {
        try
        {
            body();
        }
        catch
        {
            SettleAfterFailure();
            throw;
        }

        Settle();
    }, TestContext.Current.CancellationToken);

    public Task Run(Func<Task> body) => session.Dispatch(async () =>
    {
        try
        {
            await body();
        }
        catch
        {
            SettleAfterFailure();
            throw;
        }

        Settle();
        return true;
    }, TestContext.Current.CancellationToken);

    private static void SettleAfterFailure()
    {
        try
        {
            Settle();
        }
        catch (Exception)
        {
            // The test's own failure is the one to report.
        }
    }

    // Runs what is queued and renders the frame it asks for; what that queues in turn, once more.
    private static void Settle()
    {
        for (var i = 0; i < 2; i++)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }

        Dispatcher.UIThread.RunJobs();
    }
}

public sealed class TestApp : Application
{
    public override void Initialize() => Styles.Add(new FluentTheme());

    public static AppBuilder BuildAvaloniaApp() => AppBuilder.Configure<TestApp>()
        .UseSkia()
        .UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = false })
        .WithInterFont()
        .With(new FontManagerOptions { DefaultFamilyName = "fonts:Inter#Inter" });
}
