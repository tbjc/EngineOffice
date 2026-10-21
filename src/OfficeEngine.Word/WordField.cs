namespace OfficeEngine.Word;

/// <summary>A field written into a document, such as a page number.</summary>
/// <remarks>
/// The library writes the field; it never calculates it. The cached result is the text that a reader
/// shows until the engine that opens the document recalculates the field.
/// </remarks>
public sealed class WordField
{
    internal WordField(string instruction, string? cachedResult)
    {
        Instruction = instruction;
        CachedResult = cachedResult;
    }

    /// <summary>Gets the instruction of the field, for example <c>PAGE</c>.</summary>
    public string Instruction { get; }

    /// <summary>
    /// Gets the result stored with the field, or <see langword="null"/> when the field has none.
    /// </summary>
    public string? CachedResult { get; }

    /// <inheritdoc />
    public override string ToString() => Instruction;
}
