namespace OfficeEngine.Excel;

/// <summary>A note of a worksheet: the cell it is attached to, its text and its author.</summary>
public sealed class ExcelComment
{
    internal ExcelComment(string reference, string text, string author)
    {
        Reference = reference;
        Text = text;
        Author = author;
    }

    /// <summary>Gets the A1-style reference of the cell the note is attached to, for example <c>A1</c>.</summary>
    public string Reference { get; }

    /// <summary>Gets the text of the note.</summary>
    public string Text { get; }

    /// <summary>Gets the author of the note.</summary>
    public string Author { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Reference}: {Text}";
}
