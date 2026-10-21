using System.Globalization;

namespace OfficeEngine.Excel;

/// <summary>
/// Converts between A1-style cell references and 1-based column/row indices.
/// </summary>
internal static class CellReference
{
    /// <summary>Maximum number of columns in an .xlsx worksheet (column XFD).</summary>
    internal const int MaxColumnIndex = 16384;

    /// <summary>Maximum number of rows in an .xlsx worksheet.</summary>
    internal const int MaxRowIndex = 1048576;

    /// <summary>Parses an A1-style reference such as <c>B7</c>.</summary>
    /// <param name="reference">The reference to parse.</param>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <returns><see langword="true"/> when the reference is valid.</returns>
    internal static bool TryParse(string reference, out int column, out int row)
    {
        column = 0;
        row = 0;

        if (string.IsNullOrEmpty(reference))
        {
            return false;
        }

        int i = 0;
        int parsedColumn = 0;
        while (i < reference.Length && char.IsAsciiLetter(reference[i]))
        {
            char c = char.ToUpperInvariant(reference[i]);
            parsedColumn = (parsedColumn * 26) + (c - 'A' + 1);
            if (parsedColumn > MaxColumnIndex)
            {
                return false;
            }

            i++;
        }

        if (parsedColumn == 0 || i == 0 || i >= reference.Length)
        {
            return false;
        }

        int digitStart = i;
        long parsedRow = 0;
        while (i < reference.Length && char.IsAsciiDigit(reference[i]))
        {
            parsedRow = (parsedRow * 10) + (reference[i] - '0');
            if (parsedRow > MaxRowIndex)
            {
                return false;
            }

            i++;
        }

        if (i != reference.Length || digitStart == i || parsedRow < 1)
        {
            return false;
        }

        column = parsedColumn;
        row = (int)parsedRow;
        return true;
    }

    /// <summary>Parses an A1-style range such as <c>B2:D5</c>.</summary>
    /// <param name="range">The range reference. Both corners are required, separated by a colon.</param>
    /// <param name="firstColumn">The 1-based column index of the top-left corner.</param>
    /// <param name="firstRow">The 1-based row index of the top-left corner.</param>
    /// <param name="lastColumn">The 1-based column index of the bottom-right corner.</param>
    /// <param name="lastRow">The 1-based row index of the bottom-right corner.</param>
    /// <returns>
    /// <see langword="true"/> when the reference is valid and its corners are in top-left to
    /// bottom-right order.
    /// </returns>
    internal static bool TryParseRange(
        string range,
        out int firstColumn,
        out int firstRow,
        out int lastColumn,
        out int lastRow)
    {
        firstColumn = 0;
        firstRow = 0;
        lastColumn = 0;
        lastRow = 0;

        int separator = range.IndexOf(':');
        if (separator <= 0 || separator == range.Length - 1)
        {
            return false;
        }

        if (range.IndexOf(':', separator + 1) >= 0)
        {
            return false;
        }

        if (!TryParse(range[..separator], out firstColumn, out firstRow)
            || !TryParse(range[(separator + 1)..], out lastColumn, out lastRow))
        {
            return false;
        }

        return firstRow <= lastRow && firstColumn <= lastColumn;
    }

    /// <summary>
    /// Parses an A1-style range such as <c>B2:D5</c> or a single cell such as <c>B2</c>.
    /// </summary>
    /// <param name="reference">The range or cell reference.</param>
    /// <param name="firstColumn">The 1-based column index of the top-left corner.</param>
    /// <param name="firstRow">The 1-based row index of the top-left corner.</param>
    /// <param name="lastColumn">The 1-based column index of the bottom-right corner.</param>
    /// <param name="lastRow">The 1-based row index of the bottom-right corner.</param>
    /// <returns><see langword="true"/> when the reference is valid.</returns>
    internal static bool TryParseRangeOrCell(
        string reference,
        out int firstColumn,
        out int firstRow,
        out int lastColumn,
        out int lastRow)
    {
        if (TryParseRange(reference, out firstColumn, out firstRow, out lastColumn, out lastRow))
        {
            return true;
        }

        if (TryParse(reference, out int column, out int row))
        {
            firstColumn = column;
            firstRow = row;
            lastColumn = column;
            lastRow = row;
            return true;
        }

        firstColumn = 0;
        firstRow = 0;
        lastColumn = 0;
        lastRow = 0;
        return false;
    }

    /// <summary>Validates an A1-style range or cell and returns its normalized text.</summary>
    /// <param name="reference">The range or cell reference.</param>
    /// <param name="parameterName">The name of the parameter to report when the reference is invalid.</param>
    /// <returns>The normalized reference, for example <c>A1:B2</c>.</returns>
    internal static string NormalizeRangeOrCell(string reference, string parameterName)
    {
        if (!TryParseRangeOrCell(reference, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow))
        {
            throw new ArgumentException(
                $"'{reference}' is not a valid A1-style range reference.",
                parameterName);
        }

        string first = ToReference(firstColumn, firstRow);
        if (firstColumn == lastColumn && firstRow == lastRow)
        {
            return first;
        }

        return first + ":" + ToReference(lastColumn, lastRow);
    }

    /// <summary>Builds an A1-style reference from 1-based indices.</summary>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="row">The 1-based row index.</param>
    /// <returns>The reference, for example <c>B7</c>.</returns>
    internal static string ToReference(int column, int row)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(column, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(column, MaxColumnIndex);
        ArgumentOutOfRangeException.ThrowIfLessThan(row, 1);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(row, MaxRowIndex);

        Span<char> letters = stackalloc char[8];
        int length = 0;
        int value = column;
        while (value > 0)
        {
            int remainder = (value - 1) % 26;
            letters[length++] = (char)('A' + remainder);
            value = (value - 1) / 26;
        }

        char[] reversed = new char[length];
        for (int i = 0; i < length; i++)
        {
            reversed[i] = letters[length - 1 - i];
        }

        return string.Concat(reversed) + row.ToString(CultureInfo.InvariantCulture);
    }
}
