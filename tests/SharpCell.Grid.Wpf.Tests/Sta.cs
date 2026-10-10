using System.Runtime.ExceptionServices;

namespace SharpCell.Grid.Wpf.Tests;

/// <summary>Runs a test body on a thread of its own in a single-threaded apartment, as WPF and the clipboard need.</summary>
internal static class Sta
{
    public static void Run(Action body)
    {
        ExceptionDispatchInfo? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                body();
            }
            catch (Exception e)
            {
                failure = ExceptionDispatchInfo.Capture(e);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        failure?.Throw();
    }
}
