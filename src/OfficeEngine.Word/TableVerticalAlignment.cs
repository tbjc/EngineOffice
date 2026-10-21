namespace OfficeEngine.Word;

/// <summary>
/// Vertical alignment of the content of a table cell, mirroring OOXML <c>ST_VerticalJc</c>.
/// </summary>
public enum TableVerticalAlignment
{
    /// <summary>Content is aligned to the top of the cell (OOXML <c>top</c>).</summary>
    Top,

    /// <summary>Content is centered vertically in the cell (OOXML <c>center</c>).</summary>
    Center,

    /// <summary>Content is aligned to the bottom of the cell (OOXML <c>bottom</c>).</summary>
    Bottom,
}
