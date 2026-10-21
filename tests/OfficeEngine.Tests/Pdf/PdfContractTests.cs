using OfficeEngine;
using OfficeEngine.Pdf;

namespace OfficeEngine.Tests.Pdf;

/// <summary>The contracts of the PDF subsystem, which do not need an external engine.</summary>
public sealed class PdfContractTests
{
    [Fact]
    public void Options_RejectInvalidLimits()
    {
        PdfConversionOptions options = new() { MaximumInputBytes = 0 };
        Assert.Throws<ArgumentOutOfRangeException>(options.Validate);

        options = new PdfConversionOptions { MaximumOutputBytes = 0 };
        Assert.Throws<ArgumentOutOfRangeException>(options.Validate);

        options = new PdfConversionOptions { Timeout = TimeSpan.Zero };
        Assert.Throws<ArgumentOutOfRangeException>(options.Validate);

        options = new PdfConversionOptions { Timeout = TimeSpan.FromHours(2) };
        Assert.Throws<ArgumentOutOfRangeException>(options.Validate);
    }

    [Fact]
    public void Options_AcceptTheDefaults()
    {
        PdfConversionOptions options = new();

        options.Validate();

        Assert.Equal(64L * 1024 * 1024, options.MaximumInputBytes);
        Assert.Equal(64L * 1024 * 1024, options.MaximumOutputBytes);
        Assert.Equal(TimeSpan.FromMinutes(2), options.Timeout);
        Assert.False(options.KeepTemporaryFiles);
    }

    [Fact]
    public void Result_ExposesThePdfAndTheDiagnostics()
    {
        byte[] pdf = [1, 2, 3];

        PdfConversionResult result = new(pdf, ["mensaje"], TimeSpan.FromSeconds(2));

        Assert.Same(pdf, result.Pdf);
        Assert.Equal(["mensaje"], result.Diagnostics);
        Assert.Equal(TimeSpan.FromSeconds(2), result.Duration);
    }

    [Fact]
    public void Exception_ExposesTheExitCodeAndTheDiagnostics()
    {
        PdfConversionException exception = new("fallo", 3, ["mensaje"]);

        Assert.Equal(3, exception.ExitCode);
        Assert.Equal(["mensaje"], exception.Diagnostics);
        Assert.IsAssignableFrom<OfficeEngineException>(exception);
    }
}
