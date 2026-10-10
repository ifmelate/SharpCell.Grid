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
    private readonly Dictionary<GridColor, IImmutableSolidColorBrush> brushes = new();
    private readonly Dictionary<(GridColor, double, GridDash), ImmutablePen> pens = new();

    public AvaloniaGridCanvas(DrawingContext context)
    {
        this.context = context;
    }

    public void FillRectangle(GridRect rect, GridColor color) => context.FillRectangle(Brush(color), ToRect(rect));

    public void StrokeRectangle(GridRect rect, GridColor color, double thickness) =>
        context.DrawRectangle(null, Pen(color, thickness, GridDash.Solid), ToRect(rect));

    public void DrawLine(double x1, double y1, double x2, double y2, GridColor color, double thickness = 1, GridDash dash = GridDash.Solid) =>
        context.DrawLine(Pen(color, thickness, dash), new Point(x1, y1), new Point(x2, y2));

    public void DrawText(string text, GridRect bounds, GridTextAlignment alignment, GridVerticalAlignment verticalAlignment, GridColor color, GridFont font, bool wrap)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0)
            return;

        var formatted = Format(text, font, Brush(color));
        if (wrap)
        {
            formatted.MaxTextWidth = bounds.Width;
            formatted.TextAlignment = alignment switch
            {
                GridTextAlignment.Left => TextAlignment.Left,
                GridTextAlignment.Right => TextAlignment.Right,
                _ => TextAlignment.Center,
            };
        }

        var width = wrap ? bounds.Width : formatted.Width;
        var x = wrap ? bounds.X : alignment switch
        {
            GridTextAlignment.Left => bounds.X,
            GridTextAlignment.Right => bounds.Right - width,
            _ => bounds.X + (bounds.Width - width) / 2,
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

    /// <summary>Lays out one line of text in a font, for drawing or for measuring.</summary>
    public static FormattedText Format(string text, GridFont font, IBrush? brush)
    {
        var typeface = new Typeface(new FontFamily(font.Family), font.Italic ? FontStyle.Italic : FontStyle.Normal, font.Bold ? FontWeight.Bold : FontWeight.Normal);
        var formatted = new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, typeface, font.Size, brush);
        if (font.Underline || font.Strikethrough)
        {
            var decorations = new TextDecorationCollection();
            if (font.Underline)
                decorations.AddRange(TextDecorations.Underline);
            if (font.Strikethrough)
                decorations.AddRange(TextDecorations.Strikethrough);
            formatted.SetTextDecorations(decorations);
        }

        return formatted;
    }

    private IImmutableSolidColorBrush Brush(GridColor color)
    {
        if (!brushes.TryGetValue(color, out var brush))
            brushes[color] = brush = new ImmutableSolidColorBrush(Color.FromArgb(color.A, color.R, color.G, color.B));
        return brush;
    }

    private ImmutablePen Pen(GridColor color, double thickness, GridDash dash)
    {
        if (!pens.TryGetValue((color, thickness, dash), out var pen))
        {
            // Dash lengths are in multiples of the line's thickness.
            double[]? pattern = dash switch
            {
                GridDash.Dash => [3, 1],
                GridDash.Dot => [1, 1],
                GridDash.DashDot => [3, 1, 1, 1],
                GridDash.DashDotDot => [3, 1, 1, 1, 1, 1],
                GridDash.Hair => [1, 1],
                _ => null,
            };
            var style = pattern is null ? null : new ImmutableDashStyle(pattern, 0);
            pens[(color, thickness, dash)] = pen = new ImmutablePen(Brush(color), thickness, style, PenLineCap.Flat);
        }

        return pen;
    }

    private static Rect ToRect(GridRect rect) => new(rect.X, rect.Y, rect.Width, rect.Height);
}
