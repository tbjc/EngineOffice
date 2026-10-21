using System.Globalization;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>Parsing, validation and formatting helpers for defined names.</summary>
internal static class ExcelDefinedNames
{
    internal const int MaxNameLength = 255;

    /// <summary>Validates the text of a defined name.</summary>
    internal static void ValidateName(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        if (name.Length > MaxNameLength)
        {
            throw new ArgumentException(
                $"The defined name must be at most {MaxNameLength} characters long.",
                nameof(name));
        }

        if (!IsValidNameText(name))
        {
            throw new ArgumentException(
                "A defined name may only contain letters, digits, spaces, underscores and dots, " +
                "and must start with a letter, an underscore or a dot.",
                nameof(name));
        }

        if (CellReference.TryParse(name, out _, out _))
        {
            throw new ArgumentException(
                "A defined name cannot look like a cell reference.",
                nameof(name));
        }
    }

    /// <summary>Builds the text that a defined name stores for a range.</summary>
    internal static string FormatTarget(string? sheetName, int firstColumn, int firstRow, int lastColumn, int lastRow)
    {
        string reference = FormatAbsolute(firstColumn, firstRow, lastColumn, lastRow);
        return sheetName is null ? reference : QuoteSheetName(sheetName) + "!" + reference;
    }

    /// <summary>Formats a range as an absolute reference such as <c>$A$1:$B$2</c>.</summary>
    internal static string FormatAbsolute(int firstColumn, int firstRow, int lastColumn, int lastRow)
    {
        string first = AbsoluteCell(firstColumn, firstRow);
        if (firstColumn == lastColumn && firstRow == lastRow)
        {
            return first;
        }

        return first + ":" + AbsoluteCell(lastColumn, lastRow);
    }

    private static string AbsoluteCell(int column, int row)
    {
        string reference = CellReference.ToReference(column, row);
        int index = 0;
        while (index < reference.Length && char.IsAsciiLetter(reference[index]))
        {
            index++;
        }

        return "$" + reference[..index] + "$" + reference[index..];
    }

    /// <summary>Quotes a sheet name when the schema requires it.</summary>
    internal static string QuoteSheetName(string name)
    {
        bool simple = name.Length > 0
            && !char.IsAsciiDigit(name[0])
            && name.All(c => char.IsAsciiLetterOrDigit(c) || c is '_' or '.');
        return simple ? name : "'" + name.Replace("'", "''") + "'";
    }

    /// <summary>
    /// Reads the range that a defined name points at. Only a single cell or a single contiguous
    /// range of a worksheet of this workbook is a range; anything else (a constant, a formula, a
    /// multi-area reference or a reference to another workbook) is not.
    /// </summary>
    internal static bool TryParseTarget(
        ExcelWorkbook workbook,
        Worksheet? scope,
        string? text,
        out string? range)
    {
        range = null;
        if (string.IsNullOrEmpty(text))
        {
            return false;
        }

        string target = text;
        string? sheetName = null;
        int separator = target.LastIndexOf('!');
        if (separator >= 0)
        {
            string prefix = target[..separator];
            target = target[(separator + 1)..];
            if (prefix.Length == 0)
            {
                return false;
            }

            if (prefix[0] == '\'')
            {
                if (prefix.Length < 2 || prefix[^1] != '\'')
                {
                    return false;
                }

                sheetName = prefix[1..^1].Replace("''", "'");
            }
            else
            {
                sheetName = prefix;
            }
        }

        if (!TryParseAbsoluteRange(target, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow))
        {
            return false;
        }

        if (sheetName is null)
        {
            if (scope is null)
            {
                return false;
            }
        }
        else
        {
            Worksheet? targetSheet = workbook.FindWorksheet(sheetName);
            if (targetSheet is null)
            {
                return false;
            }

            if (scope is not null && !string.Equals(scope.Name, targetSheet.Name, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        range = firstColumn == lastColumn && firstRow == lastRow
            ? CellReference.ToReference(firstColumn, firstRow)
            : $"{CellReference.ToReference(firstColumn, firstRow)}:{CellReference.ToReference(lastColumn, lastRow)}";
        return true;
    }

    /// <summary>Parses an absolute range such as <c>$A$1</c> or <c>$A$1:$B$2</c>.</summary>
    private static bool TryParseAbsoluteRange(
        string text,
        out int firstColumn,
        out int firstRow,
        out int lastColumn,
        out int lastRow)
    {
        firstColumn = 0;
        firstRow = 0;
        lastColumn = 0;
        lastRow = 0;

        int separator = text.IndexOf(':');
        if (separator < 0)
        {
            if (!TryParseAbsoluteCell(text, out firstColumn, out firstRow))
            {
                return false;
            }

            lastColumn = firstColumn;
            lastRow = firstRow;
            return true;
        }

        if (separator == 0 || separator == text.Length - 1 || text.IndexOf(':', separator + 1) >= 0)
        {
            return false;
        }

        return TryParseAbsoluteCell(text[..separator], out firstColumn, out firstRow)
            && TryParseAbsoluteCell(text[(separator + 1)..], out lastColumn, out lastRow);
    }

    private static bool TryParseAbsoluteCell(string text, out int column, out int row)
    {
        column = 0;
        row = 0;
        int index = 0;
        if (index < text.Length && text[index] == '$')
        {
            index++;
        }

        int letterStart = index;
        while (index < text.Length && char.IsAsciiLetter(text[index]))
        {
            index++;
        }

        if (index == letterStart)
        {
            return false;
        }

        int letterEnd = index;
        if (index >= text.Length || text[index] != '$')
        {
            return false;
        }

        index++;
        int digitStart = index;
        while (index < text.Length && char.IsAsciiDigit(text[index]))
        {
            index++;
        }

        if (index == digitStart || index != text.Length)
        {
            return false;
        }

        long parsedColumn = 0;
        for (int i = letterStart; i < letterEnd; i++)
        {
            parsedColumn = (parsedColumn * 26) + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            if (parsedColumn > CellReference.MaxColumnIndex)
            {
                return false;
            }
        }

        if (!long.TryParse(
                text[digitStart..],
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out long parsedRow)
            || parsedRow < 1
            || parsedRow > CellReference.MaxRowIndex)
        {
            return false;
        }

        column = (int)parsedColumn;
        row = (int)parsedRow;
        return true;
    }

    private static bool IsValidNameText(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        char first = name[0];
        if (!char.IsAsciiLetter(first) && first is not ('_' or '.'))
        {
            return false;
        }

        foreach (char character in name)
        {
            if (!char.IsAsciiLetterOrDigit(character) && character is not ('_' or '.' or ' '))
            {
                return false;
            }
        }

        return true;
    }
}
