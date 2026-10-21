namespace OfficeEngine.Word;

/// <summary>
/// The kind of content an Office Open XML style definition applies to (<c>ST_Style</c>).
/// </summary>
/// <remarks>
/// The members mirror the four values Office Open XML defines, so any style a Word document can
/// contain maps to exactly one member.
/// </remarks>
public enum WordStyleType
{
    /// <summary>A paragraph style (<c>paragraph</c>).</summary>
    Paragraph,

    /// <summary>A character (run) style (<c>character</c>).</summary>
    Character,

    /// <summary>A table style (<c>table</c>).</summary>
    Table,

    /// <summary>A numbering style (<c>numbering</c>).</summary>
    Numbering,
}
