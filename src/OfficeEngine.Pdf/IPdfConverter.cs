namespace OfficeEngine.Pdf;

/// <summary>
/// Converts a Word document or an Excel workbook to PDF.
/// </summary>
/// <remarks>
/// <para>
/// The converter works on the bytes of the document, so the caller decides what is converted: to
/// export the <em>current</em> state of a document, save it to a stream and hand that stream over.
/// No copy of the document on disk is reused.
/// </para>
/// <para>
/// The engine that renders the PDF is an implementation detail of the adapter and never leaks into
/// this contract.
/// </para>
/// </remarks>
public interface IPdfConverter
{
    /// <summary>Converts the document read from a stream into PDF.</summary>
    /// <param name="source">
    /// The stream with the document. It is read from its current position and is not closed.
    /// </param>
    /// <param name="format">The kind of document the stream holds.</param>
    /// <param name="options">The limits and the timeout of the conversion. The defaults are used when it is not given.</param>
    /// <returns>The PDF and the diagnostics of the engine.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is empty, is not readable or is larger than the allowed input size.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="format"/> is not a defined value.</exception>
    /// <exception cref="PdfConversionException">The engine is not available or the conversion failed.</exception>
    PdfConversionResult Convert(Stream source, PdfSourceFormat format, PdfConversionOptions? options = null);
}
