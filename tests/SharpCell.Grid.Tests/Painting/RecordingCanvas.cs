namespace SharpCell.Grid.Tests.Painting;

/// <summary>A canvas that keeps what is painted on it, for asserting on draw calls.</summary>
internal sealed class RecordingCanvas : IGridCanvas
{
    public List<object> Calls { get; } = [];

    public IEnumerable<Text> Texts => Calls.OfType<Text>();

    public void FillRectangle(GridRect rect, GridColor color) => Calls.Add(new Fill(rect, color));

    public void StrokeRectangle(GridRect rect, GridColor color, double thickness) =>
        Calls.Add(new Stroke(rect, color, thickness));

    public void DrawLine(double x1, double y1, double x2, double y2, GridColor color, double thickness = 1, GridDash dash = GridDash.Solid) =>
        Calls.Add(new Line(x1, y1, x2, y2, color, thickness, dash));

    public void DrawText(string text, GridRect bounds, GridTextAlignment alignment, GridVerticalAlignment verticalAlignment, GridColor color, GridFont font, bool wrap) =>
        Calls.Add(new Text(text, bounds, alignment, verticalAlignment, color, font, wrap));

    public Text TextOf(string text) => Assert.Single(Texts, t => t.Value == text);

    /// <summary>
    /// What the selection border encloses, from the two-pixel lines drawn over the pixel lines
    /// around it; the lines under highlighted headers are left out. Null when none is drawn.
    /// </summary>
    public GridRect? Outline(GridPainter painter)
    {
        var (underX, underY) = (Math.Floor(painter.RowHeaderWidth) - 1, Math.Floor(painter.ColumnHeaderHeight) - 1);
        var lines = Calls.OfType<Line>()
            .Where(l => l.Color == painter.Theme.ActiveCellBorder && l.Thickness == 2)
            .Where(l => !(l.X1 == l.X2 && l.X1 == underX) && !(l.Y1 == l.Y2 && l.Y1 == underY))
            .ToList();
        var vertical = lines.Where(l => l.X1 == l.X2).ToList();
        var horizontal = lines.Where(l => l.Y1 == l.Y2).ToList();
        if (vertical.Count == 0 || horizontal.Count == 0)
            return null;
        var (left, right) = (vertical.Min(l => l.X1), vertical.Max(l => l.X1));
        var (top, bottom) = (horizontal.Min(l => l.Y1), horizontal.Max(l => l.Y1));
        return new GridRect(left + 0.5, top + 0.5, right - left, bottom - top);
    }

    public sealed record Fill(GridRect Rect, GridColor Color);

    public sealed record Stroke(GridRect Rect, GridColor Color, double Thickness);

    public sealed record Line(double X1, double Y1, double X2, double Y2, GridColor Color, double Thickness, GridDash Dash);

    public sealed record Text(string Value, GridRect Bounds, GridTextAlignment Alignment, GridVerticalAlignment VerticalAlignment, GridColor Color, GridFont Font, bool Wrap);
}
