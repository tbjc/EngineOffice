namespace OfficeEngine.Pdf;

/// <summary>
/// Thrown when a document cannot be converted to PDF.
/// </summary>
/// <remarks>
/// The engine that failed is not part of this contract: the message and the diagnostics say what
/// went wrong without exposing the internals of the adapter.
/// </remarks>
public sealed class PdfConversionException : OfficeEngineException
{
    /// <summary>Initializes a new instance of the <see cref="PdfConversionException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="exitCode">The exit code of the engine, when it produced one.</param>
    /// <param name="diagnostics">The messages the engine wrote while converting.</param>
    public PdfConversionException(
        string message,
        int? exitCode,
        IReadOnlyList<string> diagnostics)
        : base(message)
    {
        ExitCode = exitCode;
        Diagnostics = diagnostics;
    }

    /// <summary>Initializes a new instance of the <see cref="PdfConversionException"/> class.</summary>
    /// <param name="message">The message that describes the error.</param>
    /// <param name="exitCode">The exit code of the engine, when it produced one.</param>
    /// <param name="diagnostics">The messages the engine wrote while converting.</param>
    /// <param name="innerException">The exception that caused the current exception.</param>
    public PdfConversionException(
        string message,
        int? exitCode,
        IReadOnlyList<string> diagnostics,
        Exception innerException)
        : base(message, innerException)
    {
        ExitCode = exitCode;
        Diagnostics = diagnostics;
    }

    /// <summary>Gets the exit code of the engine, when it produced one.</summary>
    public int? ExitCode { get; }

    /// <summary>Gets the messages the engine wrote while converting.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
}
