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

    public sealed record Fill(GridRect Rect, GridColor Color);

    public sealed record Stroke(GridRect Rect, GridColor Color, double Thickness);

    public sealed record Line(double X1, double Y1, double X2, double Y2, GridColor Color, double Thickness, GridDash Dash);

    public sealed record Text(string Value, GridRect Bounds, GridTextAlignment Alignment, GridVerticalAlignment VerticalAlignment, GridColor Color, GridFont Font, bool Wrap);
}
