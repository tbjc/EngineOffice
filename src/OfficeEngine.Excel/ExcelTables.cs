using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>Builds the SpreadsheetML representation of a structured table.</summary>
internal static class ExcelTables
{
    internal const int MaxNameLength = 255;

    internal static void ValidateName(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        if (name.Length > MaxNameLength)
        {
            throw new ArgumentException(
                $"The table name must be at most {MaxNameLength} characters long.",
                nameof(name));
        }

        if (!IsValidNameText(name))
        {
            throw new ArgumentException(
                "A table name may only contain letters, digits, underscores and dots, and must start " +
                "with a letter or an underscore.",
                nameof(name));
        }

        if (CellReference.TryParse(name, out _, out _))
        {
            throw new ArgumentException("A table name cannot look like a cell reference.", nameof(name));
        }
    }

    internal static S.Table ToElement(
        string name,
        uint id,
        string range,
        bool hasHeaderRow,
        IReadOnlyList<string> columnNames)
    {
        S.Table table = new()
        {
            Id = id,
            Name = name,
            DisplayName = name,
            Reference = range,
            HeaderRowCount = hasHeaderRow ? 1U : 0U,
        };

        table.AppendChild(new S.AutoFilter { Reference = range });

        S.TableColumns columns = new() { Count = (uint)columnNames.Count };
        for (int index = 0; index < columnNames.Count; index++)
        {
            columns.AppendChild(new S.TableColumn
            {
                Id = (uint)(index + 1),
                Name = columnNames[index],
            });
        }

        table.AppendChild(columns);
        table.AppendChild(new S.TableStyleInfo
        {
            Name = "TableStyleMedium2",
            ShowFirstColumn = false,
            ShowLastColumn = false,
            ShowRowStripes = true,
            ShowColumnStripes = false,
        });

        return table;
    }

    internal static ExcelTable FromElement(S.Table table)
    {
        string name = table.DisplayName?.Value ?? table.Name?.Value ?? string.Empty;
        string range = table.Reference?.Value ?? string.Empty;
        bool hasHeaderRow = (table.HeaderRowCount?.Value ?? 1U) > 0U;
        List<string> columnNames = [];
        if (table.TableColumns is not null)
        {
            foreach (S.TableColumn column in table.TableColumns.Elements<S.TableColumn>())
            {
                columnNames.Add(column.Name?.Value ?? string.Empty);
            }
        }

        return new ExcelTable(name, range, hasHeaderRow, columnNames);
    }

    private static bool IsValidNameText(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        char first = name[0];
        if (!char.IsAsciiLetter(first) && first != '_')
        {
            return false;
        }

        foreach (char character in name)
        {
            if (!char.IsAsciiLetterOrDigit(character) && character is not ('_' or '.'))
            {
                return false;
            }
        }

        return true;
    }
}
