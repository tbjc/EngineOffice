namespace OfficeEngine.Word;

/// <summary>
/// The part a cell plays in a vertical merge, mirroring Office Open XML <c>ST_Merge</c>.
/// </summary>
/// <remarks>
/// <para>
/// A vertical merge is stored in <c>w:vMerge</c>: the topmost cell of the merged area carries
/// <c>w:val="restart"</c>, and the cells below it carry <c>w:val="continue"</c>, which Office Open
/// XML also expresses by leaving <c>w:val</c> out. A cell that is not part of a merge stores no
/// <c>w:vMerge</c> at all, which is what <see langword="null"/> means in
/// <see cref="TableCell.VerticalMerge"/>.
/// </para>
/// <para>
/// The merge only says how the cells are drawn: it does not move or hide content. A continuation
/// cell that holds text still holds it.
/// </para>
/// </remarks>
public enum TableVerticalMerge
{
    /// <summary>
    /// The cell continues a vertical merge that starts in a row above it (OOXML <c>continue</c>).
    /// </summary>
    Continue,

    /// <summary>
    /// The cell starts a vertical merge that the cells below it continue (OOXML <c>restart</c>).
    /// </summary>
    Restart,
}
