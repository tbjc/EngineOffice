using OfficeEngine.Diagnostics;

namespace OfficeEngine.Tests.Core;

public sealed class OfficeEngineLimitsTests
{
    [Fact]
    public void Default_HasPositiveLimits()
    {
        OfficeEngineLimits limits = OfficeEngineLimits.Default;

        Assert.True(limits.MaxFileSizeBytes > 0);
        Assert.True(limits.MaxEntryCount > 0);
        Assert.True(limits.MaxTotalUncompressedSizeBytes > 0);
        Assert.True(limits.MaxEntryCompressionRatio > 0);
    }

    [Fact]
    public void EnsureValid_AcceptsDefaults() => OfficeEngineLimits.Default.EnsureValid();

    [Theory]
    [InlineData(0L)]
    [InlineData(-1L)]
    public void EnsureValid_RejectsNonPositiveFileSize(long value)
    {
        OfficeEngineLimits limits = new() { MaxFileSizeBytes = value };
        Assert.Throws<ArgumentOutOfRangeException>(() => limits.EnsureValid());
    }

    [Fact]
    public void EnsureValid_RejectsNonPositiveEntryCount()
    {
        OfficeEngineLimits limits = new() { MaxEntryCount = 0 };
        Assert.Throws<ArgumentOutOfRangeException>(() => limits.EnsureValid());
    }

    [Fact]
    public void EnsureValid_RejectsNonPositiveTotalUncompressedSize()
    {
        OfficeEngineLimits limits = new() { MaxTotalUncompressedSizeBytes = 0 };
        Assert.Throws<ArgumentOutOfRangeException>(() => limits.EnsureValid());
    }

    [Fact]
    public void EnsureValid_RejectsNonPositiveCompressionRatio()
    {
        OfficeEngineLimits limits = new() { MaxEntryCompressionRatio = 0 };
        Assert.Throws<ArgumentOutOfRangeException>(() => limits.EnsureValid());
    }

    [Fact]
    public void EnsureValid_RejectsNonPositiveXmlDepth()
    {
        OfficeEngineLimits limits = new() { MaximumXmlDepth = 0 };
        Assert.Throws<ArgumentOutOfRangeException>(() => limits.EnsureValid());
    }

    [Fact]
    public void Default_LimitsTheXmlDepth()
    {
        Assert.True(OfficeEngineLimits.Default.MaximumXmlDepth > 0);
        Assert.True(OfficeEngineLimits.Default.MaximumXmlDepth <= 1024);
    }
}

public sealed class OfficeEngineDiagnosticsTests
{
    [Fact]
    public void NewDiagnostics_IsEmpty()
    {
        OfficeEngineDiagnostics diagnostics = new();

        Assert.Equal(0, diagnostics.Count);
        Assert.False(diagnostics.HasWarnings);
        Assert.False(diagnostics.HasErrors);
        Assert.Empty(diagnostics.ToSnapshot());
    }

    [Fact]
    public void Add_ThenSnapshot_KeepsOrderAndSeverity()
    {
        OfficeEngineDiagnostics diagnostics = new();
        diagnostics.Add(DiagnosticSeverity.Info, "C1", "first");
        diagnostics.Add(new OfficeEngineDiagnostic(DiagnosticSeverity.Warning, "C2", "second"));
        diagnostics.Add(DiagnosticSeverity.Error, "C3", "third");

        IReadOnlyList<OfficeEngineDiagnostic> snapshot = diagnostics.ToSnapshot();

        Assert.Equal(3, snapshot.Count);
        Assert.Equal("C1", snapshot[0].Code);
        Assert.Equal(DiagnosticSeverity.Warning, snapshot[1].Severity);
        Assert.True(diagnostics.HasWarnings);
        Assert.True(diagnostics.HasErrors);
    }

    [Fact]
    public void ToSnapshot_IsNotAffectedByLaterAdditions()
    {
        OfficeEngineDiagnostics diagnostics = new();
        diagnostics.Add(DiagnosticSeverity.Info, "C1", "first");

        IReadOnlyList<OfficeEngineDiagnostic> snapshot = diagnostics.ToSnapshot();
        diagnostics.Add(DiagnosticSeverity.Info, "C2", "second");

        Assert.Single(snapshot);
        Assert.Equal(2, diagnostics.Count);
    }

    [Fact]
    public void Clear_RemovesEverything()
    {
        OfficeEngineDiagnostics diagnostics = new();
        diagnostics.Add(DiagnosticSeverity.Info, "C1", "first");

        diagnostics.Clear();

        Assert.Equal(0, diagnostics.Count);
    }

    [Fact]
    public void Diagnostic_RejectsEmptyCodeAndMessage()
    {
        Assert.Throws<ArgumentException>(() => new OfficeEngineDiagnostic(DiagnosticSeverity.Info, string.Empty, "m"));
        Assert.Throws<ArgumentException>(() => new OfficeEngineDiagnostic(DiagnosticSeverity.Info, "c", string.Empty));
    }

    [Fact]
    public void Diagnostic_ToString_IncludesSeverityAndCode()
    {
        OfficeEngineDiagnostic diagnostic = new(DiagnosticSeverity.Warning, "OE-0001", "something happened");

        Assert.Equal("Warning OE-0001: something happened", diagnostic.ToString());
    }

    [Fact]
    public void Add_RejectsNullDiagnostic()
    {
        OfficeEngineDiagnostics diagnostics = new();
        Assert.Throws<ArgumentNullException>(() => diagnostics.Add(null!));
    }
}
