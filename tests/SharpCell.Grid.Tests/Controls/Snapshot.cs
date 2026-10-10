using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Media.Imaging;

namespace SharpCell.Grid.Tests.Controls;

/// <summary>
/// Compares a rendered frame with an approved PNG in <c>Snapshots</c> next to the test, one per
/// operating system: Skia rasterizes glyphs differently on Linux, macOS and Windows, so the same
/// frame differs in a few percent of its pixels between them. A missing or different snapshot
/// fails the test and leaves <c>name.os.received.png</c> beside it: look at it, and when it is
/// right, rename it to <c>name.os.png</c>. Small differences in anti-aliasing are tolerated; a moved
/// line or a changed text is not.
/// </summary>
internal static class Snapshot
{
    // A pixel differs when a channel is off by more than this; the frame differs when more than
    // this share of pixels does.
    private const int ChannelTolerance = 48;
    private const double PixelTolerance = 0.002;

    public static void Match(WriteableBitmap? frame, string name, [CallerFilePath] string testFile = "")
    {
        Assert.NotNull(frame);
        var folder = Path.Combine(Path.GetDirectoryName(testFile)!, "Snapshots");
        var platform = OperatingSystem.IsLinux() ? "linux" : OperatingSystem.IsMacOS() ? "macos" : "windows";
        var approved = Path.Combine(folder, $"{name}.{platform}.png");
        var received = Path.Combine(folder, $"{name}.{platform}.received.png");

        using var stream = new MemoryStream();
        frame.Save(stream, new PngBitmapEncoderOptions());
        var png = stream.ToArray();

        var problem = File.Exists(approved) ? Compare(png, File.ReadAllBytes(approved)) : "there is no approved snapshot";
        if (problem is null)
        {
            File.Delete(received);
            return;
        }

        Directory.CreateDirectory(folder);
        File.WriteAllBytes(received, png);
        Assert.Fail($"Snapshot '{name}' on {platform}: {problem}. Review {received}; rename it to {name}.{platform}.png when it is right.");
    }

    private static string? Compare(byte[] actualPng, byte[] expectedPng)
    {
        var (actual, actualSize) = Pixels(actualPng);
        var (expected, expectedSize) = Pixels(expectedPng);
        if (actualSize != expectedSize)
            return $"size {actualSize.Width}x{actualSize.Height} instead of {expectedSize.Width}x{expectedSize.Height}";

        var differing = 0;
        for (var i = 0; i < actual.Length; i += 4)
        {
            for (var channel = 0; channel < 4; channel++)
            {
                if (Math.Abs(actual[i + channel] - expected[i + channel]) > ChannelTolerance)
                {
                    differing++;
                    break;
                }
            }
        }

        var total = actual.Length / 4;
        return differing > total * PixelTolerance ? $"{differing} of {total} pixels differ" : null;
    }

    private static (byte[] Pixels, PixelSize Size) Pixels(byte[] png)
    {
        using var bitmap = new Bitmap(new MemoryStream(png));
        var size = bitmap.PixelSize;
        var stride = size.Width * 4;
        var pixels = new byte[stride * size.Height];
        var handle = GCHandle.Alloc(pixels, GCHandleType.Pinned);
        try
        {
            bitmap.CopyPixels(new PixelRect(size), handle.AddrOfPinnedObject(), pixels.Length, stride);
        }
        finally
        {
            handle.Free();
        }
        return (pixels, size);
    }
}
