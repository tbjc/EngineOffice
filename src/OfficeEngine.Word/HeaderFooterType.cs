namespace OfficeEngine.Word;

/// <summary>
/// Which pages of a section use a header or a footer.
/// </summary>
/// <remarks>
/// Office Open XML stores this as <c>w:type</c> on <c>w:headerReference</c> and
/// <c>w:footerReference</c>: <c>default</c>, <c>first</c> and <c>even</c>. A section without a
/// reference of a given type inherits the header or footer of the previous section, which is what
/// Word calls "link to previous".
/// </remarks>
public enum HeaderFooterType
{
    /// <summary>The header or footer of every page that is not covered by a more specific type.</summary>
    Primary,

    /// <summary>
    /// The header or footer of the first page. Word displays it only when the section sets
    /// <c>w:titlePg</c> (<see cref="Section.DifferentFirstPage"/>).
    /// </summary>
    FirstPage,

    /// <summary>
    /// The header or footer of the even-numbered pages. Word displays it only when the document sets
    /// <c>w:evenAndOddHeaders</c>, which this version preserves but does not model.
    /// </summary>
    EvenPage,
}
