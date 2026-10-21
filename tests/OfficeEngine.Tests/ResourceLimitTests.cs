using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests;

/// <summary>
/// Security-relevant behaviour: untrusted inputs must be rejected with typed exceptions
/// and must be bounded by the configured limits.
/// </summary>
public sealed class ResourceLimitTests
{
    [Fact]
    public void Word_Open_InputThatIsNotAZip_ThrowsInvalidOfficeFileException()
    {
        using MemoryStream stream = new(TestData.NotAZip(), writable: false);

        InvalidOfficeFileException exception = Assert.Throws<InvalidOfficeFileException>(() => WordDocument.Open(stream));
        Assert.NotNull(exception.InnerException);
    }

    [Fact]
    public void Word_Open_ZipThatIsNotADocument_ThrowsInvalidOfficeFileException()
    {
        byte[] notADocument = TestPackage.CreateZipWithEntryCount(2);
        using MemoryStream stream = new(notADocument, writable: false);

        Assert.Throws<InvalidOfficeFileException>(() => WordDocument.Open(stream));
    }

    [Fact]
    public void Excel_Open_InputThatIsNotAZip_ThrowsInvalidOfficeFileException()
    {
        using MemoryStream stream = new(TestData.NotAZip(), writable: false);

        Assert.Throws<InvalidOfficeFileException>(() => ExcelWorkbook.Open(stream));
    }

    [Fact]
    public void Excel_Open_ZipThatIsNotAWorkbook_ThrowsInvalidOfficeFileException()
    {
        byte[] notAWorkbook = TestPackage.CreateZipWithEntryCount(2);
        using MemoryStream stream = new(notAWorkbook, writable: false);

        Assert.Throws<InvalidOfficeFileException>(() => ExcelWorkbook.Open(stream));
    }

    [Fact]
    public void Word_Open_ExceedingMaxFileSize_ThrowsResourceLimit()
    {
        using MemoryStream stream = new(TestData.NotAZip(), writable: false);
        OfficeEngineLimits limits = new() { MaxFileSizeBytes = 8 };

        Assert.Throws<OfficeResourceLimitExceededException>(() => WordDocument.Open(stream, limits: limits));
    }

    [Fact]
    public void Excel_Open_ExceedingMaxFileSize_ThrowsResourceLimit()
    {
        using MemoryStream stream = new(TestData.NotAZip(), writable: false);
        OfficeEngineLimits limits = new() { MaxFileSizeBytes = 8 };

        Assert.Throws<OfficeResourceLimitExceededException>(() => ExcelWorkbook.Open(stream, limits: limits));
    }

    [Fact]
    public void Word_Open_ExceedingMaxEntryCount_ThrowsResourceLimit()
    {
        byte[] package = TestPackage.CreateZipWithEntryCount(10);
        using MemoryStream stream = new(package, writable: false);
        OfficeEngineLimits limits = new() { MaxEntryCount = 5 };

        Assert.Throws<OfficeResourceLimitExceededException>(() => WordDocument.Open(stream, limits: limits));
    }

    [Fact]
    public void Word_Open_CompressionBomb_ThrowsResourceLimitWithDefaultLimits()
    {
        byte[] bomb = TestPackage.CreateCompressibleZip("payload.bin", 2_000_000);
        using MemoryStream stream = new(bomb, writable: false);

        Assert.Throws<OfficeResourceLimitExceededException>(() => WordDocument.Open(stream));
    }

    [Fact]
    public void Excel_Open_CompressionBomb_ThrowsResourceLimitWithDefaultLimits()
    {
        byte[] bomb = TestPackage.CreateCompressibleZip("payload.bin", 2_000_000);
        using MemoryStream stream = new(bomb, writable: false);

        Assert.Throws<OfficeResourceLimitExceededException>(() => ExcelWorkbook.Open(stream));
    }

    [Fact]
    public void Word_Open_ExceedingTotalUncompressedSize_ThrowsResourceLimit()
    {
        byte[] package = TestPackage.CreateCompressibleZip("payload.bin", 100_000);
        using MemoryStream stream = new(package, writable: false);
        OfficeEngineLimits limits = new() { MaxTotalUncompressedSizeBytes = 1000, MaxEntryCompressionRatio = double.MaxValue };

        Assert.Throws<OfficeResourceLimitExceededException>(() => WordDocument.Open(stream, limits: limits));
    }

    [Fact]
    public void Word_Open_WithInvalidLimits_ThrowsArgumentOutOfRange()
    {
        using MemoryStream stream = new(TestData.NotAZip(), writable: false);
        OfficeEngineLimits limits = new() { MaxFileSizeBytes = 0 };

        Assert.Throws<ArgumentOutOfRangeException>(() => WordDocument.Open(stream, limits: limits));
    }

    [Fact]
    public void Excel_Create_WithInvalidLimits_ThrowsArgumentOutOfRange()
    {
        OfficeEngineLimits limits = new() { MaxEntryCount = 0 };

        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelWorkbook.Create(limits));
    }

    [Fact]
    public void TypedExceptions_DeriveFromOfficeEngineException()
    {
        Assert.True(typeof(OfficeEngineException).IsAssignableFrom(typeof(InvalidOfficeFileException)));
        Assert.True(typeof(OfficeEngineException).IsAssignableFrom(typeof(OfficeResourceLimitExceededException)));
        Assert.True(typeof(OfficeEngineException).IsAssignableFrom(typeof(UnsupportedOfficeFeatureException)));
        Assert.True(typeof(Exception).IsAssignableFrom(typeof(OfficeEngineException)));
    }
}
