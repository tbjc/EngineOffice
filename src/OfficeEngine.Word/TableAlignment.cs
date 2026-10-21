namespace OfficeEngine.Word;

/// <summary>
/// Horizontal alignment of a table between the page margins, mirroring OOXML <c>ST_JcTable</c>.
/// </summary>
public enum TableAlignment
{
    /// <summary>Left-aligned table (OOXML <c>left</c>).</summary>
    Left,

    /// <summary>Centered table (OOXML <c>center</c>).</summary>
    Center,

    /// <summary>Right-aligned table (OOXML <c>right</c>).</summary>
    Right,
}
