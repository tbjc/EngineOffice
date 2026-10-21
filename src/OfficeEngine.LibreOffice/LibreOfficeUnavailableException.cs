namespace OfficeEngine.LibreOffice;

/// <summary>Thrown when the LibreOffice engine is not available or is too old to be used.</summary>
internal sealed class LibreOfficeUnavailableException : Exception
{
    internal LibreOfficeUnavailableException(string message)
        : base(message)
    {
    }
}
