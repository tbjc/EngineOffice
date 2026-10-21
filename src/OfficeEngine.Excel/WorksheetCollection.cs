using System.Collections;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>
/// The ordered collection of worksheets in an <see cref="ExcelWorkbook"/>.
/// </summary>
/// <remarks>
/// The collection works directly on the workbook: changes are reflected immediately and are
/// persisted on the next <see cref="ExcelWorkbook.Save(string)"/>. It is not independently
/// serializable and it is invalidated when the workbook is disposed.
/// </remarks>
public sealed class WorksheetCollection : IReadOnlyList<Worksheet>
{
    private readonly ExcelWorkbook _workbook;

    internal WorksheetCollection(ExcelWorkbook workbook) => _workbook = workbook;

    /// <inheritdoc />
    public int Count
    {
        get
        {
            _workbook.ThrowIfDisposed();
            return SheetElements().Count;
        }
    }

    /// <summary>Gets the worksheet at the given position.</summary>
    /// <param name="index">The zero-based position of the worksheet.</param>
    /// <returns>The worksheet.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is out of range.</exception>
    public Worksheet this[int index]
    {
        get
        {
            _workbook.ThrowIfDisposed();
            List<S.Sheet> sheets = SheetElements();
            if (index < 0 || index >= sheets.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return new Worksheet(_workbook, sheets[index]);
        }
    }

    /// <summary>Gets the worksheet with the given name (case-insensitive).</summary>
    /// <param name="name">The worksheet name.</param>
    /// <returns>The worksheet.</returns>
    /// <exception cref="ArgumentException">No worksheet has the given name.</exception>
    public Worksheet this[string name]
    {
        get
        {
            ArgumentException.ThrowIfNullOrEmpty(name);
            _workbook.ThrowIfDisposed();

            foreach (S.Sheet sheet in SheetElements())
            {
                if (string.Equals(sheet.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
                {
                    return new Worksheet(_workbook, sheet);
                }
            }

            throw new ArgumentException($"The workbook does not contain a worksheet named '{name}'.", nameof(name));
        }
    }

    /// <summary>Adds a new, empty worksheet at the end of the workbook.</summary>
    /// <param name="name">The worksheet name. Must be unique (case-insensitive), at most 31 characters and free of <c>\ / ? * [ ] :</c>.</param>
    /// <returns>The new worksheet.</returns>
    /// <exception cref="ArgumentException"><paramref name="name"/> is empty, too long, contains invalid characters or is not unique.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public Worksheet Add(string name)
    {
        _workbook.ThrowIfDisposed();
        SheetName.Validate(name);
        SheetName.EnsureUnique(_workbook, name);

        WorksheetPart worksheetPart = _workbook.WorkbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());

        S.Sheet sheet = new()
        {
            Name = name,
            SheetId = NextSheetId(),
            Id = _workbook.WorkbookPart.GetIdOfPart(worksheetPart),
        };

        _workbook.Sheets().AppendChild(sheet);
        return new Worksheet(_workbook, sheet);
    }

    /// <summary>Removes a worksheet by name.</summary>
    /// <param name="name">The name of the worksheet to remove (case-insensitive).</param>
    /// <returns><see langword="true"/> when a worksheet was removed.</returns>
    /// <exception cref="InvalidOperationException">Removing the worksheet would leave the workbook without worksheets.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool Remove(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        _workbook.ThrowIfDisposed();

        List<S.Sheet> sheets = SheetElements();
        S.Sheet? sheet = sheets.FirstOrDefault(s => string.Equals(s.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
        if (sheet is null)
        {
            return false;
        }

        if (sheets.Count == 1)
        {
            throw new InvalidOperationException("A workbook must contain at least one worksheet.");
        }

        string? relationshipId = sheet.Id?.Value;
        uint index = (uint)sheets.IndexOf(sheet);

        // A name scoped to a worksheet is identified by its position in the sheet collection, so the
        // scope of the remaining names has to be moved when a sheet is removed.
        _workbook.RemoveLocalNamesOfSheet(index);
        sheet.Remove();
        if (relationshipId is not null)
        {
            _workbook.WorkbookPart.DeletePart(relationshipId);
        }

        return true;
    }

    /// <inheritdoc />
    public IEnumerator<Worksheet> GetEnumerator()
    {
        _workbook.ThrowIfDisposed();
        foreach (S.Sheet sheet in SheetElements())
        {
            yield return new Worksheet(_workbook, sheet);
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    private uint NextSheetId()
    {
        uint max = 0;
        foreach (S.Sheet sheet in SheetElements())
        {
            if (sheet.SheetId?.Value > max)
            {
                max = sheet.SheetId.Value;
            }
        }

        return max + 1;
    }

    private List<S.Sheet> SheetElements()
        => _workbook.Sheets().Elements<S.Sheet>().ToList();
}
