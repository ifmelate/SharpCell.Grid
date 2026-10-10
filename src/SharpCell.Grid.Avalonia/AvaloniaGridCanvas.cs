using System;
using System.Collections.Generic;
using System.Globalization;
using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Immutable;

namespace SharpCell.Grid.Avalonia;

/// <summary>Paints a grid through an Avalonia drawing context.</summary>
internal sealed class AvaloniaGridCanvas : IGridCanvas
{
    private readonly DrawingContext context;
    private readonly Typeface typeface;
    private readonly double fontSize;
    private readonly Dictionary<GridColor, IImmutableSolidColorBrush> brushes = new();
    private readonly Dictionary<(GridColor, double), ImmutablePen> pens = new();

    public AvaloniaGridCanvas(DrawingContext context, Typeface typeface, double fontSize)
    {
        this.context = context;
        this.typeface = typeface;
        this.fontSize = fontSize;
    }

    public void FillRectangle(GridRect rect, GridColor color) => context.FillRectangle(Brush(color), ToRect(rect));

    public void StrokeRectangle(GridRect rect, GridColor color, double thickness) =>
        context.DrawRectangle(null, Pen(color, thickness), ToRect(rect));

    public void DrawLine(double x1, double y1, double x2, double y2, GridColor color) =>
        context.DrawLine(Pen(color, 1), new Point(x1, y1), new Point(x2, y2));

    public void DrawText(string text, GridRect bounds, GridTextAlignment alignment, GridVerticalAlignment verticalAlignment, GridColor color)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0)
            return;

        var formatted = new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, typeface, fontSize, Brush(color));
        var x = alignment switch
        {
            GridTextAlignment.Left => bounds.X,
            GridTextAlignment.Right => bounds.Right - formatted.Width,
            _ => bounds.X + (bounds.Width - formatted.Width) / 2,
        };
        var y = verticalAlignment switch
        {
            GridVerticalAlignment.Top => bounds.Y,
            GridVerticalAlignment.Bottom => bounds.Bottom - formatted.Height,
            _ => bounds.Y + (bounds.Height - formatted.Height) / 2,
        };

        // Whole pixels keep glyphs crisp and frames the same from run to run.
        using (context.PushClip(ToRect(bounds)))
            context.DrawText(formatted, new Point(Math.Round(x), Math.Round(y)));
    }

    private IImmutableSolidColorBrush Brush(GridColor color)
    {
        if (!brushes.TryGetValue(color, out var brush))
            brushes[color] = brush = new ImmutableSolidColorBrush(Color.FromArgb(color.A, color.R, color.G, color.B));
        return brush;
    }

    private ImmutablePen Pen(GridColor color, double thickness)
    {
        if (!pens.TryGetValue((color, thickness), out var pen))
            pens[(color, thickness)] = pen = new ImmutablePen(Brush(color), thickness);
        return pen;
    }

    private static Rect ToRect(GridRect rect) => new(rect.X, rect.Y, rect.Width, rect.Height);
}
