namespace OfficeEngine.Tests.Excel;

/// <summary>Groups the tests that measure something and must not run in parallel with others.</summary>
[CollectionDefinition("Performance", DisableParallelization = true)]
public sealed class PerformanceGroup
{
}
