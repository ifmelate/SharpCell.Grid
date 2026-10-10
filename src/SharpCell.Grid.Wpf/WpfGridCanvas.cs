using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows;
using System.Windows.Media;

namespace SharpCell.Grid.Wpf;

/// <summary>Paints a grid through a WPF drawing context.</summary>
internal sealed class WpfGridCanvas : IGridCanvas
{
    private readonly DrawingContext context;
    private readonly double pixelsPerDip;
    private readonly Dictionary<GridColor, SolidColorBrush> brushes = new();
    private readonly Dictionary<(GridColor, double, GridDash), Pen> pens = new();

    public WpfGridCanvas(DrawingContext context, double pixelsPerDip)
    {
        this.context = context;
        this.pixelsPerDip = pixelsPerDip;
    }

    public void FillRectangle(GridRect rect, GridColor color) => context.DrawRectangle(Brush(color), null, ToRect(rect));

    public void StrokeRectangle(GridRect rect, GridColor color, double thickness) =>
        context.DrawRectangle(null, Pen(color, thickness, GridDash.Solid), ToRect(rect));

    public void DrawLine(double x1, double y1, double x2, double y2, GridColor color, double thickness = 1, GridDash dash = GridDash.Solid) =>
        context.DrawLine(Pen(color, thickness, dash), new Point(x1, y1), new Point(x2, y2));

    public void DrawText(string text, GridRect bounds, GridTextAlignment alignment, GridVerticalAlignment verticalAlignment, GridColor color, GridFont font, bool wrap)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0)
            return;

        var formatted = Format(text, font, Brush(color), pixelsPerDip);
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

        // Whole pixels keep glyphs crisp.
        context.PushClip(new RectangleGeometry(ToRect(bounds)));
        context.DrawText(formatted, new Point(Math.Round(x), Math.Round(y)));
        context.Pop();
    }

    /// <summary>Lays out text in a font, for drawing or for measuring.</summary>
    public static FormattedText Format(string text, GridFont font, Brush brush, double pixelsPerDip)
    {
        var typeface = new Typeface(new FontFamily(font.Family), font.Italic ? FontStyles.Italic : FontStyles.Normal, font.Bold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal);
        var formatted = new FormattedText(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight, typeface, font.Size, brush, pixelsPerDip);
        if (font.Underline || font.Strikethrough)
        {
            var decorations = new TextDecorationCollection();
            if (font.Underline)
                decorations.Add(TextDecorations.Underline);
            if (font.Strikethrough)
                decorations.Add(TextDecorations.Strikethrough);
            formatted.SetTextDecorations(decorations);
        }

        return formatted;
    }

    private SolidColorBrush Brush(GridColor color)
    {
        if (!brushes.TryGetValue(color, out var brush))
        {
            brush = new SolidColorBrush(Color.FromArgb(color.A, color.R, color.G, color.B));
            brush.Freeze();
            brushes[color] = brush;
        }

        return brush;
    }

    private Pen Pen(GridColor color, double thickness, GridDash dash)
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
            pen = new Pen(Brush(color), thickness)
            {
                StartLineCap = PenLineCap.Flat,
                EndLineCap = PenLineCap.Flat,
                DashCap = PenLineCap.Flat,
                DashStyle = pattern is null ? DashStyles.Solid : new DashStyle(pattern, 0),
            };
            pen.Freeze();
            pens[(color, thickness, dash)] = pen;
        }

        return pen;
    }

    private static Rect ToRect(GridRect rect) => new(rect.X, rect.Y, Math.Max(0, rect.Width), Math.Max(0, rect.Height));
}
