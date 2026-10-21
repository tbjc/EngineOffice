namespace OfficeEngine.Word;

/// <summary>
/// Plain text content of a run (<c>w:r/w:t</c>).
/// </summary>
/// <remarks>
/// Tabs and line breaks are content of their own (<see cref="RunTab"/> and <see cref="RunBreak"/>)
/// and are never part of a text item.
/// </remarks>
public sealed class RunText : RunContent
{
    internal RunText(string value) => Value = value;

    /// <summary>Gets the text.</summary>
    public string Value { get; }

    /// <inheritdoc />
    public override string ToString() => Value;
}
