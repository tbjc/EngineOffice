using OfficeEngine.Excel;

namespace OfficeEngine.Tests.Excel;

/// <summary>
/// Keeps the cost of writing many cells linear. The collection is not parallelised so that the
/// measurement of allocated memory belongs to this test alone.
/// </summary>
[Collection("Performance")]
public sealed class ExcelPerformanceTests
{
    [Fact]
    public void WritingManyCells_DoesNotAllocateProportionallyToTheSquare()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        long before = GC.GetTotalAllocatedBytes(precise: true);
        for (int row = 1; row <= 3000; row++)
        {
            sheet[row, 1].Value = $"Texto único {row}";
        }

        long allocated = GC.GetTotalAllocatedBytes(precise: true) - before;

        // Looking a string up by scanning the whole table allocated hundreds of megabytes.
        Assert.True(
            allocated < 250L * 1024 * 1024,
            $"Writing 3000 cells allocated {allocated / 1024 / 1024} MB, which suggests a quadratic lookup.");
    }
}
