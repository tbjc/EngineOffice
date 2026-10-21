using OfficeEngine.Excel;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Support;

/// <summary>
/// Saves documents to an in-memory buffer so tests can inspect the produced package.
/// </summary>
internal static class SaveBytes
{
    /// <summary>Saves a document to a new byte array.</summary>
    internal static byte[] ToBytes(WordDocument document)
    {
        using MemoryStream stream = new();
        document.Save(stream);
        return stream.ToArray();
    }

    /// <summary>Saves a workbook to a new byte array.</summary>
    internal static byte[] ToBytes(ExcelWorkbook workbook)
    {
        using MemoryStream stream = new();
        workbook.Save(stream);
        return stream.ToArray();
    }
}
