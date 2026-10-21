namespace OfficeEngine.Word;

/// <summary>
/// An explicit line break inside a run (<c>w:r/w:br</c>).
/// </summary>
/// <remarks>
/// <para>
/// Office Open XML also defines a carriage return (<c>w:r/w:cr</c>), which is a line break with no
/// attributes. Since it carries nothing this version would lose, it is reported as
/// <see cref="BreakType.TextWrapping"/> and its markup is preserved unchanged when the run is not
/// rewritten.
/// </para>
/// <para>
/// The <c>w:clear</c> attribute of <c>w:br</c>, which controls text flow around floating objects, is
/// not modelled: it is preserved on save but not reported or written.
/// </para>
/// </remarks>
public sealed class RunBreak : RunContent
{
    internal RunBreak(BreakType type) => Type = type;

    /// <summary>Gets the kind of break.</summary>
    public BreakType Type { get; }

    /// <inheritdoc />
    public override string ToString() => Type.ToString();
}
