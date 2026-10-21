namespace OfficeEngine.Word;

/// <summary>
/// How a section starts, as <c>w:sectPr/w:type</c> stores it.
/// </summary>
public enum SectionStartType
{
    /// <summary>The section starts on a new page (OOXML <c>nextPage</c>, the default).</summary>
    NextPage,

    /// <summary>The section starts on the same page (OOXML <c>continuous</c>).</summary>
    Continuous,

    /// <summary>The section starts on the next even-numbered page (OOXML <c>evenPage</c>).</summary>
    EvenPage,

    /// <summary>The section starts on the next odd-numbered page (OOXML <c>oddPage</c>).</summary>
    OddPage,

    /// <summary>The section starts on the next column (OOXML <c>nextColumn</c>).</summary>
    NextColumn,
}
