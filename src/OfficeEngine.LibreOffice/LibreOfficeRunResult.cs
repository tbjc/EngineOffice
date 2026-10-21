namespace OfficeEngine.LibreOffice;

/// <summary>What the LibreOffice process wrote and how it finished.</summary>
/// <param name="ExitCode">The exit code of the process.</param>
/// <param name="StandardOutput">Everything the process wrote to its standard output.</param>
/// <param name="StandardError">Everything the process wrote to its standard error.</param>
internal sealed record LibreOfficeRunResult(int ExitCode, string StandardOutput, string StandardError)
{
    /// <summary>Gets the messages of the process, one per line and without the empty ones.</summary>
    internal IReadOnlyList<string> Messages
    {
        get
        {
            List<string> messages = [];
            messages.AddRange(Lines(StandardOutput));
            messages.AddRange(Lines(StandardError));
            return messages;
        }
    }

    private static IEnumerable<string> Lines(string text)
    {
        foreach (string line in text.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            string trimmed = line.Trim('\r', ' ', '\t');
            if (trimmed.Length > 0)
            {
                yield return trimmed;
            }
        }
    }
}
