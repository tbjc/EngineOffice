using OfficeEngine.Formulas;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeEngine.Diagnostics;
using OfficeEngine.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>
/// A SpreadsheetML (.xlsx) workbook that can be created, opened, edited and saved.
/// </summary>
/// <remarks>
/// <para>
/// A <see cref="ExcelWorkbook"/> keeps the underlying OOXML package in memory for its whole
/// lifetime. Worksheet parts and elements that this version does not model are preserved when
/// the workbook is saved; they are reported through <see cref="Diagnostics"/> instead of being
/// dropped silently.
/// </para>
/// <para>
/// All I/O in this version is synchronous. This is a deliberate decision recorded in
/// <c>docs/adr/0004-synchronous-io-and-stream-lifecycle.md</c>.
/// </para>
/// <para>
/// This type is thread-affine: instances are not safe for concurrent use.
/// </para>
/// </remarks>
public sealed class ExcelWorkbook : IDisposable
{
    private readonly SpreadsheetDocument _package;
    private readonly MemoryStream _backing;
    private readonly OfficeEngineDiagnostics _diagnostics = new();
    private readonly WorksheetCollection _worksheets;
    private readonly Dictionary<string, int> _sharedStringIndex = new(StringComparer.Ordinal);
    private bool _disposed;

    private ExcelWorkbook(SpreadsheetDocument package, MemoryStream backing, OfficeEngineLimits limits)
    {
        _package = package;
        _backing = backing;
        Limits = limits;
        _worksheets = new WorksheetCollection(this);
    }

    /// <summary>Gets the resource limits applied when the workbook was loaded.</summary>
    public OfficeEngineLimits Limits { get; }

    /// <summary>Gets the diagnostics reported while the workbook was loaded.</summary>
    /// <remarks>The list is a snapshot; it is not updated by later edits.</remarks>
    public IReadOnlyList<OfficeEngineDiagnostic> Diagnostics => _diagnostics.ToSnapshot();

    /// <summary>Evaluates a formula over the values of this workbook.</summary>
    /// <param name="formula">The formula, with or without the leading <c>=</c>.</param>
    /// <param name="sheetName">
    /// The sheet the formula belongs to, used by the references without a sheet. The first sheet is
    /// used when it is not given.
    /// </param>
    /// <returns>
    /// The value of the formula: a number, a text, a boolean, an error of the engine or
    /// <see langword="null"/> for an empty value.
    /// </returns>
    /// <exception cref="ArgumentException"><paramref name="formula"/> is empty or too long.</exception>
    /// <exception cref="OfficeEngine.Formulas.FormulaSyntaxException">The formula cannot be read.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public object? EvaluateFormula(string formula, string? sheetName = null)
    {
        ThrowIfDisposed();
        return new FormulaEngine().Evaluate(formula, new ExcelFormulaContext(this, sheetName));
    }

    /// <summary>Gets the worksheets of the workbook.</summary>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public WorksheetCollection Worksheets
    {
        get
        {
            ThrowIfDisposed();
            return _worksheets;
        }
    }

    /// <summary>Gets the defined names of the workbook, in the order they are stored.</summary>
    /// <remarks>
    /// The catalogue includes names scoped to the whole workbook and names scoped to a worksheet, and
    /// it includes the names that this version does not interpret, such as the built-in print areas.
    /// They are preserved when the workbook is saved.
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public IReadOnlyList<ExcelDefinedName> DefinedNames
    {
        get
        {
            ThrowIfDisposed();
            return ReadDefinedNames();
        }
    }

    /// <summary>Finds a defined name by its text, ignoring case.</summary>
    /// <param name="name">The name of the defined name.</param>
    /// <returns>The defined name, or <see langword="null"/> when the workbook has no such name.</returns>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelDefinedName? FindDefinedName(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        ThrowIfDisposed();
        foreach (ExcelDefinedName definedName in ReadDefinedNames())
        {
            if (string.Equals(definedName.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                return definedName;
            }
        }

        return null;
    }

    /// <summary>Removes a defined name that is scoped to the whole workbook.</summary>
    /// <param name="name">The name of the defined name.</param>
    /// <returns><see langword="true"/> when a name was removed.</returns>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool RemoveNamedRange(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        ThrowIfDisposed();
        return RemoveDefinedName(name, localSheetId: null);
    }

    /// <summary>Creates a new workbook in memory with a single worksheet named <c>Sheet1</c>.</summary>
    /// <param name="limits">Optional resource limits. Defaults to <see cref="OfficeEngineLimits.Default"/>.</param>
    /// <returns>A new workbook.</returns>
    public static ExcelWorkbook Create(OfficeEngineLimits? limits = null)
    {
        OfficeEngineLimits effective = limits ?? OfficeEngineLimits.Default;
        effective.EnsureValid();

        MemoryStream backing = new();
        SpreadsheetDocument? package = null;
        try
        {
            package = SpreadsheetDocument.Create(backing, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = package.AddWorkbookPart();
            workbookPart.Workbook = new S.Workbook(new S.Sheets());

            ExcelWorkbook workbook = new(package, backing, effective);
            workbook._worksheets.Add("Sheet1");
            return workbook;
        }
        catch
        {
            package?.Dispose();
            backing.Dispose();
            throw;
        }
    }

    /// <summary>Opens a workbook from a file.</summary>
    /// <param name="path">Path of the .xlsx file.</param>
    /// <param name="limits">Optional resource limits. Defaults to <see cref="OfficeEngineLimits.Default"/>.</param>
    /// <returns>The opened workbook.</returns>
    /// <exception cref="InvalidOfficeFileException">The file is not a readable SpreadsheetML workbook.</exception>
    public static ExcelWorkbook Open(string path, OfficeEngineLimits? limits = null)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);
        OfficeEngineLimits effective = limits ?? OfficeEngineLimits.Default;
        MemoryStream backing = OoxmlLoader.LoadFromFile(path, effective);
        return OpenCore(backing, effective);
    }

    /// <summary>Opens a workbook from a stream.</summary>
    /// <param name="stream">The stream that contains the .xlsx package.</param>
    /// <param name="leaveOpen">
    /// When <see langword="false"/> (the default), <paramref name="stream"/> is disposed after it has been read.
    /// The workbook never holds a reference to the passed stream.
    /// </param>
    /// <param name="limits">Optional resource limits. Defaults to <see cref="OfficeEngineLimits.Default"/>.</param>
    /// <returns>The opened workbook.</returns>
    /// <exception cref="InvalidOfficeFileException">The stream is not a readable SpreadsheetML workbook.</exception>
    public static ExcelWorkbook Open(Stream stream, bool leaveOpen = false, OfficeEngineLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        OfficeEngineLimits effective = limits ?? OfficeEngineLimits.Default;
        MemoryStream backing = OoxmlLoader.LoadFromStream(stream, leaveOpen, effective);
        return OpenCore(backing, effective);
    }

    /// <summary>Saves the workbook to a file, replacing any existing file atomically.</summary>
    /// <param name="path">Destination path.</param>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public void Save(string path)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrEmpty(path);

        string fullPath = Path.GetFullPath(path);
        string directory = Path.GetDirectoryName(fullPath) ?? ".";
        string temporaryPath = Path.Combine(directory, $"{Path.GetRandomFileName()}.tmp");
        try
        {
            using (FileStream file = new(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                Save(file);
            }

            File.Move(temporaryPath, fullPath, overwrite: true);
        }
        catch
        {
            TryDelete(temporaryPath);
            throw;
        }
    }

    /// <summary>Saves the workbook to a stream.</summary>
    /// <param name="destination">The writable destination stream. It is left open.</param>
    /// <exception cref="ArgumentNullException"><paramref name="destination"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public void Save(Stream destination)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(destination);
        if (!destination.CanWrite)
        {
            throw new ArgumentException("The destination stream must be writable.", nameof(destination));
        }

        UpdateSharedStringCounts();
        _package.Save();
        byte[] bytes = _backing.ToArray();
        destination.Write(bytes, 0, bytes.Length);
    }
    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _package.Dispose();
        _backing.Dispose();
    }

    internal WorkbookPart WorkbookPart
        => _package.WorkbookPart
           ?? throw new InvalidOfficeFileException("The package does not contain a workbook part, so it is not a SpreadsheetML workbook.");

    internal S.Sheets Sheets()
    {
        S.Workbook workbook = OoxmlPartReader.Read("the workbook part", () => WorkbookPart.Workbook);
        return workbook.Sheets
            ?? throw new InvalidOfficeFileException("The workbook part does not contain a sheets element.");
    }

    internal bool Uses1904DateSystem
        => WorkbookPart.Workbook?.WorkbookProperties?.Date1904?.Value ?? false;

    internal Worksheet? FindWorksheet(string name)
    {
        foreach (S.Sheet sheet in Sheets().Elements<S.Sheet>())
        {
            if (string.Equals(sheet.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
            {
                return new Worksheet(this, sheet);
            }
        }

        return null;
    }

    internal Worksheet? WorksheetAt(uint index)
    {
        List<S.Sheet> sheets = Sheets().Elements<S.Sheet>().ToList();
        return index < sheets.Count ? new Worksheet(this, sheets[(int)index]) : null;
    }

    internal uint SheetIndexOf(S.Sheet sheet)
    {
        uint index = 0;
        foreach (S.Sheet candidate in Sheets().Elements<S.Sheet>())
        {
            if (ReferenceEquals(candidate, sheet))
            {
                return index;
            }

            index++;
        }

        throw new InvalidOfficeFileException("The worksheet does not belong to this workbook.");
    }

    internal S.DefinedNames GetOrCreateDefinedNames()
    {
        S.Workbook root = WorkbookPart.Workbook
            ?? throw new InvalidOfficeFileException("The package does not contain a readable workbook part.");
        S.DefinedNames? names = root.GetFirstChild<S.DefinedNames>();
        if (names is null)
        {
            names = new S.DefinedNames();
            ExcelWorkbookLayout.InsertInOrder(root, names);
        }

        return names;
    }

    /// <summary>
    /// Removes the names that were scoped to a worksheet that is being removed and moves the scope of
    /// the names that came after it, so that the local scope keeps pointing at the same worksheet.
    /// </summary>
    internal void RemoveLocalNamesOfSheet(uint index)
    {
        S.DefinedNames? names = WorkbookPart.Workbook?.GetFirstChild<S.DefinedNames>();
        if (names is null)
        {
            return;
        }

        foreach (S.DefinedName element in names.Elements<S.DefinedName>().ToList())
        {
            uint? local = element.LocalSheetId?.Value;
            if (local is null)
            {
                continue;
            }

            if (local.Value == index)
            {
                element.Remove();
            }
            else if (local.Value > index)
            {
                element.LocalSheetId = local.Value - 1;
            }
        }
    }

    internal uint NextTableId()
    {
        uint max = 0;
        foreach (S.Sheet sheet in Sheets().Elements<S.Sheet>())
        {
            if (sheet.Id?.Value is not string relationshipId
                || WorkbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts)
            {
                uint id = tablePart.Table?.Id?.Value ?? 0U;
                if (id > max)
                {
                    max = id;
                }
            }
        }

        return max + 1;
    }

    internal void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    /// <summary>
    /// Returns the index of a string in the shared string table, adding it when it is not there yet.
    /// The same text is stored only once.
    /// </summary>
    internal int GetOrAddSharedString(string text)
    {
        SharedStringTablePart? part = WorkbookPart.SharedStringTablePart;
        if (part is null)
        {
            part = WorkbookPart.AddNewPart<SharedStringTablePart>();
        }

        S.SharedStringTable table = OoxmlPartReader.ReadOrDefault(
            "the shared string table part",
            () => part.SharedStringTable) ?? CreateSharedStringTable(part);

        // The index avoids scanning the table, and with it the quadratic cost of writing many
        // cells: the text of every entry would be rebuilt on every lookup.
        if (!_sharedStringIndex.TryGetValue(text, out int index))
        {
            index = table.Elements<S.SharedStringItem>().Count();
            table.AppendChild(new S.SharedStringItem(
                new S.Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            _sharedStringIndex[text] = index;
        }

        SharedStringReferences++;
        return index;
    }

    private void IndexSharedStrings()
    {
        _sharedStringIndex.Clear();
        SharedStringTablePart? part = WorkbookPart.SharedStringTablePart;
        S.SharedStringTable? table = OoxmlPartReader.ReadOrDefault(
            "the shared string table part",
            () => part?.SharedStringTable);
        if (table is null)
        {
            return;
        }

        int index = 0;
        foreach (S.SharedStringItem item in table.Elements<S.SharedStringItem>())
        {
            _sharedStringIndex.TryAdd(item.InnerText, index);
            index++;
        }
    }

    /// <summary>The number of cells that refer to the shared string table.</summary>
    internal int SharedStringReferences { get; set; }

    internal string? GetSharedString(int index)
    {
        if (index < 0)
        {
            return null;
        }

        SharedStringTablePart? part = WorkbookPart.SharedStringTablePart;
        S.SharedStringTable? table = OoxmlPartReader.ReadOrDefault(
            "the shared string table part",
            () => part?.SharedStringTable);
        if (table is null)
        {
            return null;
        }

        S.SharedStringItem? item = table.Elements<S.SharedStringItem>().ElementAtOrDefault(index);
        return item?.InnerText;
    }

    private List<ExcelDefinedName> ReadDefinedNames()
    {
        S.DefinedNames? names = WorkbookPart.Workbook?.GetFirstChild<S.DefinedNames>();
        if (names is null)
        {
            return [];
        }

        List<ExcelDefinedName> result = [];
        foreach (S.DefinedName element in names.Elements<S.DefinedName>())
        {
            Worksheet? scope = element.LocalSheetId?.Value is uint localSheetId
                ? WorksheetAt(localSheetId)
                : null;
            result.Add(new ExcelDefinedName(this, element, scope));
        }

        return result;
    }

    internal bool RemoveDefinedName(string name, uint? localSheetId)
    {
        S.DefinedNames? names = WorkbookPart.Workbook?.GetFirstChild<S.DefinedNames>();
        if (names is null)
        {
            return false;
        }

        S.DefinedName? target = names.Elements<S.DefinedName>().FirstOrDefault(element =>
            string.Equals(element.Name?.Value, name, StringComparison.OrdinalIgnoreCase)
            && element.LocalSheetId?.Value == localSheetId);
        if (target is null)
        {
            return false;
        }

        target.Remove();
        if (!names.Elements<S.DefinedName>().Any())
        {
            names.Remove();
        }

        return true;
    }

    private void UpdateSharedStringCounts()
    {
        S.SharedStringTable? table = OoxmlPartReader.ReadOrDefault(
            "the shared string table part",
            () => WorkbookPart.SharedStringTablePart?.SharedStringTable);
        if (table is null)
        {
            return;
        }

        table.Count = (uint)Math.Max(0, SharedStringReferences);
        table.UniqueCount = (uint)table.Elements<S.SharedStringItem>().Count();
    }

    private static S.SharedStringTable CreateSharedStringTable(SharedStringTablePart part)
    {
        S.SharedStringTable table = new();
        part.SharedStringTable = table;
        return table;
    }

    private static ExcelWorkbook OpenCore(MemoryStream backing, OfficeEngineLimits limits)
    {
        SpreadsheetDocument? package = null;
        try
        {
            OoxmlPackageInspector.EnsureWithinLimits(backing, limits);
            OoxmlXmlGuard.EnsureSafe(backing, limits);
            backing.Position = 0;

            try
            {
                package = SpreadsheetDocument.Open(backing, isEditable: true);
                WorkbookPart workbookPart = package.WorkbookPart
                    ?? throw new InvalidOfficeFileException("The package does not contain a workbook part, so it is not a SpreadsheetML workbook.");
                workbookPart.Workbook ??= new S.Workbook();
                workbookPart.Workbook.Sheets ??= new S.Sheets();
            }
            catch (Exception ex) when (ex is OpenXmlPackageException or FileFormatException
                or InvalidDataException or InvalidOperationException or XmlException or ArgumentOutOfRangeException)
            {
                throw new InvalidOfficeFileException("The input is not a readable SpreadsheetML (.xlsx) workbook.", ex);
            }

            ExcelWorkbook workbook = new(package, backing, limits);
            workbook.SharedStringReferences = (int)(OoxmlPartReader.ReadOrDefault(
                "the shared string table part",
                () => package.WorkbookPart?.SharedStringTablePart?.SharedStringTable)?.Count?.Value ?? 0U);
            workbook.IndexSharedStrings();
            workbook.Inspect();
            return workbook;
        }
        catch
        {
            package?.Dispose();
            backing.Dispose();
            throw;
        }
    }

    private void Inspect()
    {
        try
        {
            InspectCore();
        }
        catch (Exception ex) when (ex is InvalidDataException or XmlException or OpenXmlPackageException
            or FileFormatException or InvalidOperationException or ArgumentOutOfRangeException)
        {
            throw new InvalidOfficeFileException("The package cannot be read as a SpreadsheetML (.xlsx) workbook.", ex);
        }
    }

    private void InspectCore()
    {
        foreach (S.Sheet sheet in Sheets().Elements<S.Sheet>())
        {
            if (sheet.Id?.Value is not string relationshipId)
            {
                continue;
            }

            if (WorkbookPart.GetPartById(relationshipId) is not WorksheetPart worksheetPart)
            {
                continue;
            }

            S.Worksheet worksheet = OoxmlPartReader.Read(
                $"the worksheet '{sheet.Name?.Value}'",
                () => worksheetPart.Worksheet);

            List<string> elementNames = [];
            int unmodelled = 0;
            foreach (DocumentFormat.OpenXml.OpenXmlElement child in worksheet.ChildElements)
            {
                if (child.LocalName is "sheetData" or "dataValidations" or "autoFilter" or "tableParts"
                    or "cols" or "sheetViews")
                {
                    continue;
                }

                unmodelled++;
                if (!elementNames.Contains(child.LocalName))
                {
                    elementNames.Add(child.LocalName);
                }
            }

            if (unmodelled > 0)
            {
                _diagnostics.Add(
                    DiagnosticSeverity.Info,
                    ExcelDiagnosticCodes.UnmodelledWorksheetContent,
                    $"Worksheet '{sheet.Name?.Value}' contains {unmodelled} element(s) that this version does not model " +
                    $"({string.Join(", ", elementNames)}). They are preserved when the workbook is saved but cannot be edited through the API.");
            }
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
            // Best effort cleanup; the original exception is more important.
        }
        catch (UnauthorizedAccessException)
        {
            // Best effort cleanup; the original exception is more important.
        }
    }
}
