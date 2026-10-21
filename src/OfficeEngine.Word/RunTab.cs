namespace OfficeEngine.Word;

/// <summary>
/// A tab character inside a run (<c>w:r/w:tab</c>).
/// </summary>
/// <remarks>
/// The tab advances to the next tab stop defined by the paragraph or inherited from its style or
/// the default tab interval. The position of a tab stop is modelled by <see cref="TabStop"/>.
/// </remarks>
public sealed class RunTab : RunContent
{
    internal RunTab()
    {
    }

    /// <inheritdoc />
    public override string ToString() => "\t";
}
