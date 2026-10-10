namespace SharpCell.Grid;

/// <summary>
/// What <see cref="GridPainter"/> paints on. Each UI framework implements it over its own drawing
/// API; coordinates are device-independent pixels from the grid's top-left corner.
/// </summary>
public interface IGridCanvas
{
    /// <summary>Fills a rectangle.</summary>
    void FillRectangle(GridRect rect, GridColor color);

    /// <summary>Draws the outline of a rectangle, the stroke centred on its edges.</summary>
    void StrokeRectangle(GridRect rect, GridColor color, double thickness);

    /// <summary>Draws a one-pixel line. Lines of the grid lie on pixel centres, such as <c>x = 63.5</c>.</summary>
    void DrawLine(double x1, double y1, double x2, double y2, GridColor color);

    /// <summary>Draws one line of text placed inside the bounds and cut off at them.</summary>
    void DrawText(string text, GridRect bounds, GridTextAlignment alignment, GridVerticalAlignment verticalAlignment, GridColor color);
}
