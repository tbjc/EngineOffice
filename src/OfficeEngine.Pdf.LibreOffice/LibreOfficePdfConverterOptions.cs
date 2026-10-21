namespace OfficeEngine.Pdf.LibreOffice;

/// <summary>The settings of the LibreOffice headless converter.</summary>
public sealed class LibreOfficePdfConverterOptions
{
    /// <summary>
    /// The path of the LibreOffice executable. When it is not given, the usual installation
    /// directories and the <c>PATH</c> are searched.
    /// </summary>
    public string? ExecutablePath { get; set; }

    /// <summary>
    /// The lowest version of LibreOffice that the converter accepts, as <c>major.minor</c>. The
    /// default is <c>7.0</c>. A build whose version cannot be read is accepted.
    /// </summary>
    public string MinimumVersion { get; set; } = "7.0";

    /// <summary>How long the check of the version may take. The default is thirty seconds.</summary>
    public TimeSpan VersionCheckTimeout { get; set; } = TimeSpan.FromSeconds(30);

    /// <summary>
    /// How many conversions this instance runs at the same time. The default is two, which keeps
    /// the number of running engine processes bounded without serialising every conversion.
    /// </summary>
    public int MaximumConcurrentConversions { get; set; } = 2;
}
