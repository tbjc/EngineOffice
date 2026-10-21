namespace OfficeEngine.Word;

/// <summary>
/// The justification of a list level's numbering (<c>w:lvlJc</c>).
/// </summary>
/// <remarks>
/// Office Open XML restricts <c>w:lvlJc</c> to these three values, unlike the paragraph alignment
/// <c>ST_Jc</c> it is named after. A stored value outside this set cannot be read and is reported as
/// <see cref="InvalidOfficeFileException"/>.
/// </remarks>
public enum WordListJustification
{
    /// <summary>The numbering is aligned to the left of its space (<c>left</c>).</summary>
    Left,

    /// <summary>The numbering is centered in its space (<c>center</c>).</summary>
    Center,

    /// <summary>The numbering is aligned to the right of its space (<c>right</c>).</summary>
    Right,
}
