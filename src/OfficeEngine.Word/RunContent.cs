namespace OfficeEngine.Word;

/// <summary>
/// One modelled piece of the content of a <see cref="Run"/>.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="Run.Content"/> returns the parts of a run that this version models, in document
/// order: text (<see cref="RunText"/>), tab characters (<see cref="RunTab"/>) and line breaks
/// (<see cref="RunBreak"/>). This is the faithful view of a run; <see cref="Run.Text"/> only
/// concatenates plain text.
/// </para>
/// <para>
/// Content that Office Open XML defines but this version does not model (drawings, fields,
/// absolute-position tabs, <c>w:noBreakHyphen</c>, <c>w:lastRenderedPageBreak</c>…) is preserved
/// across a save but is not represented here, so consecutive text items may have been separated by
/// markup that is not returned.
/// </para>
/// </remarks>
public abstract class RunContent
{
    private protected RunContent()
    {
    }
}
