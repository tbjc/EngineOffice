namespace OfficeEngine.Word;

/// <summary>A place where a searched text was found.</summary>
public sealed class WordTextMatch
{
    internal WordTextMatch(Paragraph paragraph, int startIndex, int length, string text)
    {
        Paragraph = paragraph;
        StartIndex = startIndex;
        Length = length;
        Text = text;
    }

    /// <summary>Gets the paragraph that contains the match.</summary>
    public Paragraph Paragraph { get; }

    /// <summary>Gets the offset of the match inside the text of the paragraph.</summary>
    public int StartIndex { get; }

    /// <summary>Gets the length of the match.</summary>
    public int Length { get; }

    /// <summary>Gets the text that was found.</summary>
    public string Text { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Paragraph} @ {StartIndex}: {Text}";
}
