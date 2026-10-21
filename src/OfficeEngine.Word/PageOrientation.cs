namespace OfficeEngine.Word;

/// <summary>
/// The orientation of a page, as <c>w:pgSz/@w:orient</c> stores it.
/// </summary>
public enum PageOrientation
{
    /// <summary>The page is taller than it is wide (OOXML <c>portrait</c>).</summary>
    Portrait,

    /// <summary>The page is wider than it is tall (OOXML <c>landscape</c>).</summary>
    Landscape,
}
