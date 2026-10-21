namespace OfficeEngine.Tests.Support;

/// <summary>
/// Shared, obviously-synthetic values used across tests.
/// </summary>
internal static class TestData
{
    /// <summary>Bytes that are not a ZIP container.</summary>
    internal static byte[] NotAZip() => "this is definitely not an Open XML package"u8.ToArray();

    /// <summary>Creates a unique temporary directory for a test.</summary>
    internal static string NewTemporaryDirectory()
    {
        string path = Path.Combine(Path.GetTempPath(), "OfficeEngine.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }
}
