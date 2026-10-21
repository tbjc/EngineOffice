namespace OfficeEngine.Word;

/// <summary>
/// Kind of explicit line break inside a run, mirroring OOXML <c>ST_BrType</c>.
/// </summary>
public enum BreakType
{
    /// <summary>A line break that wraps the text without starting a new page or column.</summary>
    TextWrapping,

    /// <summary>A page break.</summary>
    Page,

    /// <summary>A column break.</summary>
    Column,
}
