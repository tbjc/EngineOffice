namespace OfficeEngine.Legacy;

/// <summary>The outcome of a successful conversion of a legacy document.</summary>
public sealed class LegacyConversionResult
{
    internal LegacyConversionResult(
        byte[] document,
        LegacyDocumentFormat format,
        IReadOnlyList<string> diagnostics,
        TimeSpan duration)
    {
        Document = document;
        Format = format;
        Diagnostics = diagnostics;
        Duration = duration;
    }

    /// <summary>Gets the bytes of the converted document, in the modern format.</summary>
    public byte[] Document { get; }

    /// <summary>Gets the format of the source document.</summary>
    public LegacyDocumentFormat Format { get; }

    /// <summary>
    /// Gets the messages the engine wrote while converting. They describe the conversion, never the
    /// content of the document.
    /// </summary>
    public IReadOnlyList<string> Diagnostics { get; }

    /// <summary>Gets how long the conversion took.</summary>
    public TimeSpan Duration { get; }
}
