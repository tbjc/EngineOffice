using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// Shared access to the border edges (<c>w:top</c>, <c>w:left</c>, …) that <c>w:tblBorders</c> and
/// <c>w:tcBorders</c> both hold.
/// </summary>
/// <remarks>
/// The Open XML SDK does not reorder the children of these elements, and the schema requires a
/// sequence, so edges are inserted following <see cref="Order"/> rather than appended. A child this
/// version does not model (a diagonal edge, for example) is left where it is and never rewritten.
/// </remarks>
internal static class BorderEdges
{
    private static readonly string[] Order =
    [
        "top", "left", "bottom", "right", "insideH", "insideV", "tl2br", "tr2bl",
    ];

    /// <summary>Reads one edge of a border container.</summary>
    /// <typeparam name="TEdge">The edge element to read, for example <c>W.TopBorder</c>.</typeparam>
    /// <param name="container">The <c>w:tblBorders</c> or <c>w:tcBorders</c> element, or <see langword="null"/>.</param>
    /// <returns>The stored border, or <see langword="null"/> when the edge is absent.</returns>
    internal static TableBorder? Read<TEdge>(OpenXmlCompositeElement? container)
        where TEdge : W.BorderType
    {
        W.BorderType? edge = container?.GetFirstChild<TEdge>();
        return edge is null ? null : TableBorder.Read(edge);
    }

    /// <summary>Writes or removes one edge of a border container.</summary>
    /// <typeparam name="TEdge">The edge element to write, for example <c>W.TopBorder</c>.</typeparam>
    /// <param name="container">The <c>w:tblBorders</c> or <c>w:tcBorders</c> element to write into.</param>
    /// <param name="value">The border to store, or <see langword="null"/> to remove the edge.</param>
    internal static void Write<TEdge>(OpenXmlCompositeElement container, TableBorder? value)
        where TEdge : W.BorderType, new()
    {
        TEdge? existing = container.GetFirstChild<TEdge>();
        if (value is not { } border)
        {
            existing?.Remove();
            return;
        }

        TEdge edge = existing ?? new TEdge();
        border.ApplyTo(edge);
        if (existing is null)
        {
            InsertInOrder(container, edge);
        }
    }

    private static void InsertInOrder(OpenXmlCompositeElement container, OpenXmlElement element)
        => SchemaOrder.Insert(container, element, Order);
}
