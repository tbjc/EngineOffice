using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>
/// Validation rules for worksheet names, shared by creation and renaming.
/// </summary>
internal static class SheetName
{
    /// <summary>Validates the syntax of a worksheet name.</summary>
    /// <param name="name">The candidate name.</param>
    internal static void Validate(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        if (name.Length > 31)
        {
            throw new ArgumentException("A worksheet name cannot exceed 31 characters.", nameof(name));
        }

        if (name.IndexOfAny(['\\', '/', '?', '*', '[', ']', ':']) >= 0)
        {
            throw new ArgumentException("A worksheet name cannot contain any of the characters \\ / ? * [ ] :", nameof(name));
        }
    }

    /// <summary>Ensures no other worksheet in the workbook uses the name (case-insensitive).</summary>
    /// <param name="workbook">The workbook to inspect.</param>
    /// <param name="name">The candidate name.</param>
    /// <param name="exclude">A worksheet that is allowed to keep the name, usually the one being renamed.</param>
    internal static void EnsureUnique(ExcelWorkbook workbook, string name, S.Sheet? exclude = null)
    {
        S.Sheets sheets = workbook.Sheets();
        foreach (S.Sheet sheet in sheets.Elements<S.Sheet>())
        {
            if (ReferenceEquals(sheet, exclude))
            {
                continue;
            }

            if (string.Equals(sheet.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException($"A worksheet named '{name}' already exists.", nameof(name));
            }
        }
    }
}
