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

    /// <summary>
    /// Draws a line, centred on the coordinates. One-pixel lines of the grid lie on pixel centres,
    /// such as <c>x = 63.5</c>.
    /// </summary>
    void DrawLine(double x1, double y1, double x2, double y2, GridColor color, double thickness = 1, GridDash dash = GridDash.Solid);

    /// <summary>
    /// Draws text placed inside the bounds and cut off at them: one line, or wrapped onto as many
    /// lines as the width needs.
    /// </summary>
    void DrawText(string text, GridRect bounds, GridTextAlignment alignment, GridVerticalAlignment verticalAlignment, GridColor color, GridFont font, bool wrap);
}
