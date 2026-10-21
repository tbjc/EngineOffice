using System.Globalization;

namespace OfficeEngine.Benchmarks;

/// <summary>One measured operation.</summary>
/// <param name="Name">The name of the scenario.</param>
/// <param name="Elapsed">How long the operation took.</param>
/// <param name="AllocatedBytes">How much memory the operation allocated.</param>
/// <param name="PackageBytes">The size of the package the operation produced, when it produced one.</param>
internal sealed record BenchmarkResult(
    string Name,
    TimeSpan Elapsed,
    long AllocatedBytes,
    long PackageBytes)
{
    internal string Format()
        => string.Create(
            CultureInfo.InvariantCulture,
            $"{Name,-28} {Elapsed.TotalMilliseconds,10:0.0} ms {AllocatedBytes / 1024d / 1024d,8:0.0} MB {PackageBytes / 1024d / 1024d,8:0.0} MB");
}
