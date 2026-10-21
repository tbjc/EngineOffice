namespace OfficeEngine.Pdf;

/// <summary>The limits and the timeout that apply to one conversion.</summary>
public sealed class PdfConversionOptions
{
    /// <summary>The maximum size of the source document, in bytes. The default is 64 MiB.</summary>
    public long MaximumInputBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>The maximum size of the produced PDF, in bytes. The default is 64 MiB.</summary>
    public long MaximumOutputBytes { get; init; } = 64L * 1024 * 1024;

    /// <summary>How long the conversion may take. The default is two minutes.</summary>
    public TimeSpan Timeout { get; init; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Whether the temporary files of the conversion are kept, which is useful to diagnose a
    /// failure. The default is <see langword="false"/>.
    /// </summary>
    public bool KeepTemporaryFiles { get; init; }

    /// <summary>Checks that the limits and the timeout are usable.</summary>
    /// <exception cref="ArgumentOutOfRangeException">A limit is not positive or the timeout is out of range.</exception>
    public void Validate()
    {
        if (MaximumInputBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(MaximumInputBytes), MaximumInputBytes, "The maximum input size must be greater than zero.");
        }

        if (MaximumOutputBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(MaximumOutputBytes), MaximumOutputBytes, "The maximum output size must be greater than zero.");
        }

        if (Timeout <= TimeSpan.Zero || Timeout > TimeSpan.FromHours(1))
        {
            throw new ArgumentOutOfRangeException(
                nameof(Timeout), Timeout, "The timeout must be greater than zero and at most one hour.");
        }
    }
}
