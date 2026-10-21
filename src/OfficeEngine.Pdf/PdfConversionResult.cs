namespace OfficeEngine.Pdf;

/// <summary>The outcome of a successful conversion.</summary>
public sealed class PdfConversionResult
{
    /// <summary>Initializes a new instance of the <see cref="PdfConversionResult"/> class.</summary>
    /// <param name="pdf">The bytes of the produced PDF.</param>
    /// <param name="diagnostics">The messages the engine wrote while converting.</param>
    /// <param name="duration">How long the conversion took.</param>
    public PdfConversionResult(byte[] pdf, IReadOnlyList<string> diagnostics, TimeSpan duration)
    {
        Pdf = pdf;
        Diagnostics = diagnostics;
        Duration = duration;
    }

    /// <summary>Gets the bytes of the produced PDF.</summary>
    public byte[] Pdf { get; }

    /// <summary>
    /// Gets the messages the engine wrote while converting. They describe the conversion, never the
    /// content of the document.
    /// </summary>
    public IReadOnlyList<string> Diagnostics { get; }

    /// <summary>Gets how long the conversion took.</summary>
    public TimeSpan Duration { get; }
}
