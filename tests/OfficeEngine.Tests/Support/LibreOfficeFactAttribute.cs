using OfficeEngine.Pdf.LibreOffice;
using Xunit.Sdk;

namespace OfficeEngine.Tests.Support;

/// <summary>
/// A test that needs the external rendering engine. It is skipped, with an explicit reason, when
/// LibreOffice is not installed in the environment.
/// </summary>
public sealed class LibreOfficeFactAttribute : FactAttribute
{
    /// <summary>Initializes a new instance of the <see cref="LibreOfficeFactAttribute"/> class.</summary>
    public LibreOfficeFactAttribute()
    {
        if (!LibreOfficeTestEnvironment.IsAvailable)
        {
            Skip = "LibreOffice is not installed in this environment.";
        }
    }
}

/// <summary>Knows whether the external rendering engine can be used by the tests.</summary>
public static class LibreOfficeTestEnvironment
{
    /// <summary>Gets whether the LibreOffice executable was found.</summary>
    public static bool IsAvailable { get; } = Detect();

    private static bool Detect()
    {
        using LibreOfficePdfConverter converter = new();
        return converter.IsAvailable;
    }
}
