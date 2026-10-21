using System.Globalization;
using System.Text;

namespace OfficeEngine.Benchmarks;

/// <summary>The entry point of the benchmark harness.</summary>
internal static class Program
{
    private static int Main(string[] args)
    {
        bool update = args.Contains("--update-baseline", StringComparer.Ordinal);
        string? baselinePath = ValueOf(args, "--baseline");
        double threshold = 2.0;

        List<BenchmarkResult> results = Scenarios.Run();
        Console.WriteLine(Benchmark.Format(results));

        if (baselinePath is not null && File.Exists(baselinePath) && !update)
        {
            Dictionary<string, (double Milliseconds, long Allocated, long Package)> baseline =
                Scenarios.ParseBaseline(File.ReadAllText(baselinePath));
            int regressions = 0;
            Console.WriteLine();
            Console.WriteLine($"Comparación con {baselinePath} (umbral {threshold:0.0}x):");
            foreach (BenchmarkResult result in results)
            {
                if (!baseline.TryGetValue(result.Name, out (double Milliseconds, long Allocated, long Package) reference))
                {
                    Console.WriteLine($"  {result.Name,-28} sin referencia");
                    continue;
                }

                double factor = reference.Milliseconds <= 0
                    ? 1
                    : result.Elapsed.TotalMilliseconds / reference.Milliseconds;
                string verdict = factor > threshold ? " REGRESIÓN" : string.Empty;
                if (factor > threshold)
                {
                    regressions++;
                }

                Console.WriteLine(string.Create(
                    CultureInfo.InvariantCulture,
                    $"  {result.Name,-28} {reference.Milliseconds,8:0.0} ms -> {result.Elapsed.TotalMilliseconds,8:0.0} ms ({factor,5:0.00}x){verdict}"));
            }

            return regressions > 0 ? 1 : 0;
        }

        if (update && baselinePath is not null)
        {
            File.WriteAllText(baselinePath, Scenarios.ToBaseline(results));
            Console.WriteLine();
            Console.WriteLine($"Referencia escrita en {baselinePath}");
        }

        return 0;
    }

    private static string? ValueOf(string[] args, string name)
    {
        for (int index = 0; index < args.Length - 1; index++)
        {
            if (string.Equals(args[index], name, StringComparison.Ordinal))
            {
                return args[index + 1];
            }
        }

        return null;
    }
}
