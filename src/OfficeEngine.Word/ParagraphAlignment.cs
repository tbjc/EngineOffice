namespace OfficeEngine.Word;

/// <summary>
/// Horizontal alignment of a paragraph, mirroring a subset of OOXML <c>ST_Jc</c>.
/// </summary>
public enum ParagraphAlignment
{
    /// <summary>Left-aligned text.</summary>
    Left,

    /// <summary>Centered text.</summary>
    Center,

    /// <summary>Right-aligned text.</summary>
    Right,

    /// <summary>Justified text (distributed).</summary>
    Justify,

    /// <summary>Justified text, including the last line (equivalent to OOXML <c>both</c>).</summary>
    Both,
}