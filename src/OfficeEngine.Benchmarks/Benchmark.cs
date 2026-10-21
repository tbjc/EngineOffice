using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace OfficeEngine.Benchmarks;

/// <summary>Measures one operation and returns how long it took and how much it allocated.</summary>
internal static class Benchmark
{
    internal static BenchmarkResult Measure(string name, Action action, Func<long>? packageSize = null)
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        long before = GC.GetTotalAllocatedBytes(precise: true);
        Stopwatch watch = Stopwatch.StartNew();
        action();
        watch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - before;

        return new BenchmarkResult(name, watch.Elapsed, allocated, packageSize?.Invoke() ?? 0);
    }

    /// <summary>Writes the results as a table that a human can read and a tool can compare.</summary>
    internal static string Format(IEnumerable<BenchmarkResult> results)
    {
        StringBuilder builder = new();
        builder.AppendLine(string.Create(
            CultureInfo.InvariantCulture,
            $"{"escenario",-28} {"tiempo",13} {"asignado",13} {"paquete",13}"));
        foreach (BenchmarkResult result in results)
        {
            builder.AppendLine(result.Format());
        }

        return builder.ToString();
    }
}
