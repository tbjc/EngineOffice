namespace OfficeEngine.Legacy;

/// <summary>Thrown when a legacy document cannot be converted.</summary>
public sealed class LegacyConversionException : Exception
{
    internal LegacyConversionException(
        string message,
        int? exitCode = null,
        IReadOnlyList<string>? diagnostics = null,
        Exception? innerException = null)
        : base(message, innerException)
    {
        ExitCode = exitCode;
        Diagnostics = diagnostics ?? [];
    }

    /// <summary>Gets the exit code of the engine, when it produced one.</summary>
    public int? ExitCode { get; }

    /// <summary>Gets the messages the engine wrote while converting.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
}
