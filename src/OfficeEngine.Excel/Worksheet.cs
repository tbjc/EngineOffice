using OfficeEngine.OpenXml;
using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>
/// A live view over a single worksheet of an <see cref="ExcelWorkbook"/>.
/// </summary>
public sealed class Worksheet
{
    private const double MaxRowHeightPoints = 409d;
    private const double MaxColumnWidthCharacters = 255d;

    private readonly ExcelWorkbook _workbook;
    private readonly S.Sheet _sheet;

    internal Worksheet(ExcelWorkbook workbook, S.Sheet sheet)
    {
        _workbook = workbook;
        _sheet = sheet;
    }

    /// <summary>Gets or sets the name of the worksheet.</summary>
    /// <value>
    /// The name must be non-empty, at most 31 characters, free of <c>\ / ? * [ ] :</c> and unique
    /// within the workbook (case-insensitive).
    /// </value>
    /// <exception cref="ArgumentException">The assigned name is invalid.</exception>
    public string Name
    {
        get => _sheet.Name?.Value ?? string.Empty;
        set
        {
            SheetName.Validate(value);
            SheetName.EnsureUnique(_workbook, value, _sheet);
            _sheet.Name = value;
        }
    }

    /// <summary>
    /// Gets a cell by A1-style reference, for example <c>B7</c>.
    /// </summary>
    /// <param name="cellReference">The A1-style cell reference.</param>
    /// <returns>A cell handle. Reading a handle for a cell that does not exist yields a null value.</returns>
    /// <exception cref="ArgumentException"><paramref name="cellReference"/> is not a valid A1-style reference.</exception>
    public Cell this[string cellReference] => GetCell(cellReference);

    /// <summary>Gets a cell by A1-style reference.</summary>
    /// <param name="cellReference">The A1-style cell reference, for example <c>B7</c>.</param>
    /// <returns>A cell handle. Reading a handle for a cell that does not exist yields a null value.</returns>
    /// <exception cref="ArgumentException"><paramref name="cellReference"/> is not a valid A1-style reference.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public Cell GetCell(string cellReference)
    {
        ArgumentException.ThrowIfNullOrEmpty(cellReference);
        _workbook.ThrowIfDisposed();

        if (!CellReference.TryParse(cellReference, out int column, out int row))
        {
            throw new ArgumentException($"'{cellReference}' is not a valid A1-style cell reference.", nameof(cellReference));
        }

        return new Cell(_workbook, this, CellReference.ToReference(column, row), column, row);
    }

    /// <summary>Gets a cell by 1-based row and column indices.</summary>
    /// <param name="row">The 1-based row index (1 to 1,048,576).</param>
    /// <param name="column">The 1-based column index (1 to 16,384).</param>
    /// <returns>A cell handle. Reading a handle for a cell that does not exist yields a null value.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="row"/> or <paramref name="column"/> is outside the range of an .xlsx worksheet.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public Cell GetCell(int row, int column)
    {
        _workbook.ThrowIfDisposed();
        ValidateIndex(row, nameof(row), CellReference.MaxRowIndex);
        ValidateIndex(column, nameof(column), CellReference.MaxColumnIndex);

        return new Cell(_workbook, this, CellReference.ToReference(column, row), column, row);
    }

    /// <summary>Gets a rectangular range by an A1-style reference such as <c>B2:D5</c>.</summary>
    /// <param name="rangeReference">
    /// The range reference. Both corners are required, separated by a colon and in top-left to
    /// bottom-right order.
    /// </param>
    /// <returns>A live range over the worksheet.</returns>
    /// <exception cref="ArgumentException">
    /// <paramref name="rangeReference"/> is not a valid A1-style range or its corners are reversed.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public CellRange GetRange(string rangeReference)
    {
        ArgumentException.ThrowIfNullOrEmpty(rangeReference);
        _workbook.ThrowIfDisposed();

        if (!CellReference.TryParseRange(
            rangeReference,
            out int firstColumn,
            out int firstRow,
            out int lastColumn,
            out int lastRow))
        {
            throw new ArgumentException(
                $"'{rangeReference}' is not a valid A1-style range reference with its corners in " +
                "top-left to bottom-right order.",
                nameof(rangeReference));
        }

        return new CellRange(_workbook, this, firstRow, firstColumn, lastRow, lastColumn);
    }

    /// <summary>Gets a rectangular range by 1-based row and column indices.</summary>
    /// <param name="firstRow">The 1-based row index of the top-left corner.</param>
    /// <param name="firstColumn">The 1-based column index of the top-left corner.</param>
    /// <param name="lastRow">The 1-based row index of the bottom-right corner.</param>
    /// <param name="lastColumn">The 1-based column index of the bottom-right corner.</param>
    /// <returns>A live range over the worksheet.</returns>
    /// <exception cref="ArgumentOutOfRangeException">A corner is outside the range of an .xlsx worksheet.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="lastRow"/> is before <paramref name="firstRow"/> or
    /// <paramref name="lastColumn"/> is before <paramref name="firstColumn"/>.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public CellRange GetRange(int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        _workbook.ThrowIfDisposed();
        ValidateIndex(firstRow, nameof(firstRow), CellReference.MaxRowIndex);
        ValidateIndex(firstColumn, nameof(firstColumn), CellReference.MaxColumnIndex);
        ValidateIndex(lastRow, nameof(lastRow), CellReference.MaxRowIndex);
        ValidateIndex(lastColumn, nameof(lastColumn), CellReference.MaxColumnIndex);

        if (lastRow < firstRow)
        {
            throw new ArgumentException("The last row must not be before the first row.", nameof(lastRow));
        }

        if (lastColumn < firstColumn)
        {
            throw new ArgumentException("The last column must not be before the first column.", nameof(lastColumn));
        }

        return new CellRange(_workbook, this, firstRow, firstColumn, lastRow, lastColumn);
    }

    /// <summary>Gets a cell by 1-based row and column indices, for example <c>sheet[7, 2]</c> is <c>B7</c>.</summary>
    /// <param name="row">The 1-based row index (1 to 1,048,576).</param>
    /// <param name="column">The 1-based column index (1 to 16,384).</param>
    /// <returns>A cell handle. Reading a handle for a cell that does not exist yields a null value.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="row"/> or <paramref name="column"/> is outside the range of an .xlsx worksheet.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public Cell this[int row, int column] => GetCell(row, column);

    /// <summary>Adds a named range scoped to the whole workbook.</summary>
    /// <param name="name">
    /// The name of the range, for example <c>Datos</c>. It may contain letters, digits, spaces,
    /// underscores and dots, must start with a letter, an underscore or a dot, cannot look like a
    /// cell reference and is at most 255 characters long.
    /// </param>
    /// <param name="rangeReference">
    /// The A1-style range of this worksheet, for example <c>A1:B2</c>. A single cell such as
    /// <c>C3</c> is also accepted.
    /// </param>
    /// <returns>The defined name that was created or replaced.</returns>
    /// <remarks>
    /// A name that already exists with the same text and the same scope has its target replaced.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// The name is invalid or the range reference is not a valid A1-style reference.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelDefinedName AddNamedRange(string name, string rangeReference)
    {
        ArgumentException.ThrowIfNullOrEmpty(rangeReference);
        _workbook.ThrowIfDisposed();
        ExcelDefinedNames.ValidateName(name);
        if (!CellReference.TryParseRangeOrCell(rangeReference, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow))
        {
            throw new ArgumentException(
                $"'{rangeReference}' is not a valid A1-style range reference.",
                nameof(rangeReference));
        }

        S.DefinedNames names = _workbook.GetOrCreateDefinedNames();
        S.DefinedName? element = FindDefinedNameElement(names, name, localSheetId: null);
        if (element is null)
        {
            element = new S.DefinedName { Name = name };
            names.AppendChild(element);
        }

        element.Text = ExcelDefinedNames.FormatTarget(Name, firstColumn, firstRow, lastColumn, lastRow);
        return new ExcelDefinedName(_workbook, element, scope: null);
    }

    /// <summary>Adds a named range scoped to this worksheet.</summary>
    /// <param name="name">The name of the range, with the same rules as <see cref="AddNamedRange"/>.</param>
    /// <param name="rangeReference">
    /// The A1-style range of this worksheet, for example <c>A1:B2</c>. A single cell such as <c>C3</c>
    /// is also accepted.
    /// </param>
    /// <returns>The defined name that was created or replaced.</returns>
    /// <remarks>
    /// A name that already exists with the same text and the same scope has its target replaced. A
    /// name scoped to a worksheet is removed when the worksheet is removed.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// The name is invalid or the range reference is not a valid A1-style reference.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelDefinedName AddLocalNamedRange(string name, string rangeReference)
    {
        ArgumentException.ThrowIfNullOrEmpty(rangeReference);
        _workbook.ThrowIfDisposed();
        ExcelDefinedNames.ValidateName(name);
        if (!CellReference.TryParseRangeOrCell(rangeReference, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow))
        {
            throw new ArgumentException(
                $"'{rangeReference}' is not a valid A1-style range reference.",
                nameof(rangeReference));
        }

        uint scope = _workbook.SheetIndexOf(_sheet);
        S.DefinedNames names = _workbook.GetOrCreateDefinedNames();
        S.DefinedName? element = FindDefinedNameElement(names, name, scope);
        if (element is null)
        {
            element = new S.DefinedName { Name = name, LocalSheetId = scope };
            names.AppendChild(element);
        }

        element.Text = ExcelDefinedNames.FormatTarget(null, firstColumn, firstRow, lastColumn, lastRow);
        return new ExcelDefinedName(_workbook, element, this);
    }

    /// <summary>Removes a named range that is scoped to this worksheet.</summary>
    /// <param name="name">The name of the defined name.</param>
    /// <returns><see langword="true"/> when a name was removed.</returns>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool RemoveLocalNamedRange(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        _workbook.ThrowIfDisposed();
        return _workbook.RemoveDefinedName(name, _workbook.SheetIndexOf(_sheet));
    }

    /// <summary>Gets the data validation rules of the worksheet.</summary>
    /// <remarks>
    /// The rules are read from the worksheet each time they are asked for. A rule that applies to
    /// several areas reports the first one; the rest of the areas stay in the file.
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public IReadOnlyList<ExcelDataValidation> DataValidations
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.DataValidations? container = Part().Worksheet!.GetFirstChild<S.DataValidations>();
            if (container is null)
            {
                return [];
            }

            return container.Elements<S.DataValidation>()
                .Select(ExcelDataValidations.FromElement)
                .ToList();
        }
    }

    /// <summary>Adds a data validation rule to the worksheet.</summary>
    /// <param name="validation">The rule to add.</param>
    /// <returns>The rule that was added.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="validation"/> is <see langword="null"/>.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelDataValidation AddDataValidation(ExcelDataValidation validation)
    {
        ArgumentNullException.ThrowIfNull(validation);
        _workbook.ThrowIfDisposed();

        S.Worksheet worksheet = Part().Worksheet!;
        S.DataValidations? container = worksheet.GetFirstChild<S.DataValidations>();
        if (container is null)
        {
            container = new S.DataValidations();
            ExcelWorksheetLayout.InsertInOrder(worksheet, container);
        }

        container.AppendChild(ExcelDataValidations.ToElement(validation));
        container.Count = (uint)container.Elements<S.DataValidation>().Count();
        return validation;
    }

    /// <summary>Removes the data validation rule that applies to a range.</summary>
    /// <param name="rangeReference">The A1-style range of the rule, for example <c>A1:A10</c>.</param>
    /// <returns><see langword="true"/> when a rule was removed.</returns>
    /// <exception cref="ArgumentException"><paramref name="rangeReference"/> is not a valid A1-style reference.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool RemoveDataValidation(string rangeReference)
    {
        string range = ExcelDataValidations.NormalizeRange(rangeReference);
        _workbook.ThrowIfDisposed();

        S.Worksheet worksheet = Part().Worksheet!;
        S.DataValidations? container = worksheet.GetFirstChild<S.DataValidations>();
        if (container is null)
        {
            return false;
        }

        S.DataValidation? target = container.Elements<S.DataValidation>().FirstOrDefault(element =>
            string.Equals(
                ExcelDataValidations.FirstArea(element.SequenceOfReferences?.InnerText),
                range,
                StringComparison.OrdinalIgnoreCase));
        if (target is null)
        {
            return false;
        }

        target.Remove();
        container.Count = (uint)container.Elements<S.DataValidation>().Count();
        if (container.Count == 0U)
        {
            container.Remove();
        }

        return true;
    }

    /// <summary>Gets or sets the auto filter of the worksheet.</summary>
    /// <value>
    /// The filter and the range it covers, or <see langword="null"/> when the worksheet has no auto
    /// filter. Assigning <see langword="null"/> removes it.
    /// </value>
    /// <remarks>
    /// The filter criteria (which rows are hidden) are not modelled: they are preserved when the
    /// worksheet is saved untouched and dropped when the filter is written again.
    /// </remarks>
    /// <exception cref="InvalidOfficeFileException">The stored auto filter is malformed.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelAutoFilter? AutoFilter
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.AutoFilter? element = Part().Worksheet!.GetFirstChild<S.AutoFilter>();
            if (element is null)
            {
                return null;
            }

            string? reference = element.Reference?.Value;
            if (string.IsNullOrEmpty(reference))
            {
                throw new InvalidOfficeFileException("The worksheet has an auto filter without a range.");
            }

            return new ExcelAutoFilter(reference);
        }
        set
        {
            _workbook.ThrowIfDisposed();
            S.Worksheet worksheet = Part().Worksheet!;
            S.AutoFilter? element = worksheet.GetFirstChild<S.AutoFilter>();
            if (value is null)
            {
                element?.Remove();
                return;
            }

            if (element is null)
            {
                element = new S.AutoFilter();
                ExcelWorksheetLayout.InsertInOrder(worksheet, element);
            }

            element.RemoveAllChildren<S.FilterColumn>();
            element.Reference = value.Range;
        }
    }

    /// <summary>Gets the structured tables of the worksheet.</summary>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public IReadOnlyList<ExcelTable> Tables
    {
        get
        {
            _workbook.ThrowIfDisposed();
            List<ExcelTable> tables = [];
            foreach (TableDefinitionPart tablePart in Part().TableDefinitionParts)
            {
                if (tablePart.Table is not null)
                {
                    tables.Add(ExcelTables.FromElement(tablePart.Table));
                }
            }

            return tables;
        }
    }

    /// <summary>Adds a structured table to the worksheet.</summary>
    /// <param name="name">
    /// The name of the table, for example <c>Tabla1</c>. It may contain letters, digits, underscores
    /// and dots, must start with a letter or an underscore, cannot look like a cell reference and has
    /// to be unique in the workbook.
    /// </param>
    /// <param name="rangeReference">The A1-style range of the table, for example <c>A1:C5</c>.</param>
    /// <param name="hasHeaderRow">
    /// Whether the first row of the range is the header row. When it is, the column names are taken
    /// from the text of those cells; the empty ones get a generated name.
    /// </param>
    /// <returns>The table that was created.</returns>
    /// <exception cref="ArgumentException">
    /// The name or the range is invalid, or a table with the same name already exists in the
    /// workbook.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelTable AddTable(string name, string rangeReference, bool hasHeaderRow = true)
    {
        ArgumentException.ThrowIfNullOrEmpty(rangeReference);
        _workbook.ThrowIfDisposed();
        ExcelTables.ValidateName(name);
        if (!CellReference.TryParseRangeOrCell(
                rangeReference,
                out int firstColumn,
                out int firstRow,
                out int lastColumn,
                out int lastRow))
        {
            throw new ArgumentException(
                $"'{rangeReference}' is not a valid A1-style range reference.",
                nameof(rangeReference));
        }

        if (FindTable(name) is not null)
        {
            throw new ArgumentException(
                $"The workbook already contains a table named '{name}'.",
                nameof(name));
        }

        string range = CellReference.NormalizeRangeOrCell(rangeReference, nameof(rangeReference));
        List<string> columnNames = [];
        for (int column = firstColumn; column <= lastColumn; column++)
        {
            string? text = hasHeaderRow ? ReadCellText(column, firstRow) : null;
            string candidate = string.IsNullOrEmpty(text) ? $"Column{columnNames.Count + 1}" : text!;
            columnNames.Add(Unique(columnNames, candidate));
        }

        WorksheetPart part = Part();
        TableDefinitionPart tablePart = part.AddNewPart<TableDefinitionPart>();
        tablePart.Table = ExcelTables.ToElement(
            name,
            _workbook.NextTableId(),
            range,
            hasHeaderRow,
            columnNames);

        S.Worksheet worksheet = part.Worksheet!;
        S.TableParts? tableParts = worksheet.GetFirstChild<S.TableParts>();
        if (tableParts is null)
        {
            tableParts = new S.TableParts();
            ExcelWorksheetLayout.InsertInOrder(worksheet, tableParts);
        }

        tableParts.AppendChild(new S.TablePart { Id = part.GetIdOfPart(tablePart) });
        tableParts.Count = (uint)tableParts.Elements<S.TablePart>().Count();

        return new ExcelTable(name, range, hasHeaderRow, columnNames);
    }

    /// <summary>Removes a structured table of the worksheet.</summary>
    /// <param name="name">The name of the table.</param>
    /// <returns><see langword="true"/> when a table was removed.</returns>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool RemoveTable(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        _workbook.ThrowIfDisposed();

        WorksheetPart part = Part();
        foreach (TableDefinitionPart tablePart in part.TableDefinitionParts)
        {
            S.Table? table = tablePart.Table;
            string tableName = table?.DisplayName?.Value ?? table?.Name?.Value ?? string.Empty;
            if (!string.Equals(tableName, name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string relationshipId = part.GetIdOfPart(tablePart);
            S.TableParts? tableParts = part.Worksheet!.GetFirstChild<S.TableParts>();
            S.TablePart? reference = tableParts?.Elements<S.TablePart>()
                .FirstOrDefault(candidate => candidate.Id?.Value == relationshipId);
            reference?.Remove();
            if (tableParts is not null)
            {
                tableParts.Count = (uint)tableParts.Elements<S.TablePart>().Count();
                if (tableParts.Count == 0U)
                {
                    tableParts.Remove();
                }
            }

            part.DeletePart(tablePart);
            return true;
        }

        return false;
    }

    /// <summary>Gets or sets the frozen part of the worksheet.</summary>
    /// <value>
    /// The frozen rows and columns, or <see langword="null"/> when the sheet has no frozen panes.
    /// Assigning <see langword="null"/> removes them.
    /// </value>
    /// <exception cref="InvalidOfficeFileException">The stored pane is malformed.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelFreezePanes? FreezePanes
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.SheetView? view = FindSheetView();
            S.Pane? pane = view?.GetFirstChild<S.Pane>();
            if (pane is null || pane.State?.Value != S.PaneStateValues.Frozen)
            {
                // A split pane (not frozen) stores the split in twips, not in a number of rows and
                // columns, so it is not a frozen pane.
                return null;
            }

            int columns = (int)(pane.HorizontalSplit?.Value ?? 0d);
            int rows = (int)(pane.VerticalSplit?.Value ?? 0d);
            if (rows == 0 && columns == 0)
            {
                return null;
            }

            return new ExcelFreezePanes(rows, columns);
        }
        set
        {
            _workbook.ThrowIfDisposed();
            S.Worksheet worksheet = Part().Worksheet!;
            S.SheetView view = GetOrCreateSheetView(worksheet);
            view.RemoveAllChildren<S.Pane>();
            if (value is null)
            {
                return;
            }

            S.Pane pane = new()
            {
                TopLeftCell = value.TopLeftCell,
                ActivePane = new S.PaneValues(value.ActivePane),
                State = S.PaneStateValues.Frozen,
            };
            if (value.Columns > 0)
            {
                pane.HorizontalSplit = value.Columns;
            }

            if (value.Rows > 0)
            {
                pane.VerticalSplit = value.Rows;
            }

            view.InsertAt(pane, 0);
        }
    }

    /// <summary>Gets the height of a row in points, or <see langword="null"/> when it is the default.</summary>
    /// <param name="row">The 1-based row index.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="row"/> is out of range.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public double? RowHeight(int row)
    {
        ValidateIndex(row, nameof(row), CellReference.MaxRowIndex);
        _workbook.ThrowIfDisposed();
        return FindRow(SheetData(), row)?.Height?.Value;
    }

    /// <summary>Sets the height of a row in points.</summary>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="points">The height, greater than 0 and at most 409 points.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The row index is out of range or the height is outside the supported range.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public void SetRowHeight(int row, double points)
    {
        ValidateIndex(row, nameof(row), CellReference.MaxRowIndex);
        if (!double.IsFinite(points) || points <= 0d || points > MaxRowHeightPoints)
        {
            throw new ArgumentOutOfRangeException(
                nameof(points), points, $"The row height must be greater than 0 and at most {MaxRowHeightPoints} points.");
        }

        _workbook.ThrowIfDisposed();
        S.SheetData sheetData = SheetData();
        S.Row rowElement = FindRow(sheetData, row) ?? InsertRow(sheetData, row);
        rowElement.Height = points;
        rowElement.CustomHeight = true;
    }

    /// <summary>Gets whether a row is hidden.</summary>
    /// <param name="row">The 1-based row index.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="row"/> is out of range.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool IsRowHidden(int row)
    {
        ValidateIndex(row, nameof(row), CellReference.MaxRowIndex);
        _workbook.ThrowIfDisposed();
        return FindRow(SheetData(), row)?.Hidden?.Value ?? false;
    }

    /// <summary>Hides or shows a row.</summary>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="hidden">Whether the row is hidden.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="row"/> is out of range.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public void SetRowHidden(int row, bool hidden)
    {
        ValidateIndex(row, nameof(row), CellReference.MaxRowIndex);
        _workbook.ThrowIfDisposed();
        S.SheetData sheetData = SheetData();
        S.Row rowElement = FindRow(sheetData, row) ?? InsertRow(sheetData, row);
        rowElement.Hidden = hidden;
    }

    /// <summary>
    /// Gets the width of a column in characters of the default font, or <see langword="null"/> when it
    /// is the default.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="column"/> is out of range.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public double? ColumnWidth(int column)
    {
        ValidateIndex(column, nameof(column), CellReference.MaxColumnIndex);
        _workbook.ThrowIfDisposed();
        S.Columns? columns = Part().Worksheet!.GetFirstChild<S.Columns>();
        if (columns is null)
        {
            return null;
        }

        S.Column? element = columns.Elements<S.Column>()
            .FirstOrDefault(candidate => candidate.Min?.Value <= column && candidate.Max?.Value >= column);
        return element?.Width?.Value;
    }

    /// <summary>Sets the width of a column in characters of the default font.</summary>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="characters">The width, greater than 0 and at most 255 characters.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The column index is out of range or the width is outside the supported range.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public void SetColumnWidth(int column, double characters)
    {
        ValidateIndex(column, nameof(column), CellReference.MaxColumnIndex);
        if (!double.IsFinite(characters) || characters <= 0d || characters > MaxColumnWidthCharacters)
        {
            throw new ArgumentOutOfRangeException(
                nameof(characters),
                characters,
                $"The column width must be greater than 0 and at most {MaxColumnWidthCharacters} characters.");
        }

        _workbook.ThrowIfDisposed();
        S.Worksheet worksheet = Part().Worksheet!;
        S.Columns columns = GetOrCreateColumns(worksheet);
        S.Column element = EnsureColumn(columns, column);
        element.Width = characters;
        element.CustomWidth = true;
    }

    /// <summary>Gets whether a column is hidden.</summary>
    /// <param name="column">The 1-based column index.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="column"/> is out of range.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool IsColumnHidden(int column)
    {
        ValidateIndex(column, nameof(column), CellReference.MaxColumnIndex);
        _workbook.ThrowIfDisposed();
        S.Columns? columns = Part().Worksheet!.GetFirstChild<S.Columns>();
        if (columns is null)
        {
            return false;
        }

        S.Column? element = columns.Elements<S.Column>()
            .FirstOrDefault(candidate => candidate.Min?.Value <= column && candidate.Max?.Value >= column);
        return element?.Hidden?.Value ?? false;
    }

    /// <summary>Hides or shows a column.</summary>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="hidden">Whether the column is hidden.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="column"/> is out of range.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public void SetColumnHidden(int column, bool hidden)
    {
        ValidateIndex(column, nameof(column), CellReference.MaxColumnIndex);
        _workbook.ThrowIfDisposed();
        S.Worksheet worksheet = Part().Worksheet!;
        S.Columns columns = GetOrCreateColumns(worksheet);
        S.Column element = EnsureColumn(columns, column);
        element.Hidden = hidden;
    }

    /// <summary>Gets or sets the margins of the printed page.</summary>
    /// <value>
    /// The margins, in inches, or <see langword="null"/> when the sheet has no margins. Assigning
    /// <see langword="null"/> removes them.
    /// </value>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelPageMargins? PageMargins
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.PageMargins? element = Part().Worksheet!.GetFirstChild<S.PageMargins>();
            if (element is null)
            {
                return null;
            }

            return new ExcelPageMargins
            {
                Left = element.Left?.Value ?? 0.7,
                Right = element.Right?.Value ?? 0.7,
                Top = element.Top?.Value ?? 0.75,
                Bottom = element.Bottom?.Value ?? 0.75,
                Header = element.Header?.Value ?? 0.3,
                Footer = element.Footer?.Value ?? 0.3,
            };
        }
        set
        {
            _workbook.ThrowIfDisposed();
            S.Worksheet worksheet = Part().Worksheet!;
            S.PageMargins? element = worksheet.GetFirstChild<S.PageMargins>();
            if (value is null)
            {
                element?.Remove();
                return;
            }

            if (element is null)
            {
                element = new S.PageMargins();
                ExcelWorksheetLayout.InsertInOrder(worksheet, element);
            }

            element.Left = value.Left;
            element.Right = value.Right;
            element.Top = value.Top;
            element.Bottom = value.Bottom;
            element.Header = value.Header;
            element.Footer = value.Footer;
        }
    }

    /// <summary>Gets or sets the setup of the printed page.</summary>
    /// <value>
    /// The setup, or <see langword="null"/> when the sheet has none. Assigning
    /// <see langword="null"/> removes it.
    /// </value>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelPageSetup? PageSetup
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.PageSetup? element = Part().Worksheet!.GetFirstChild<S.PageSetup>();
            if (element is null)
            {
                return null;
            }

            return new ExcelPageSetup
            {
                Orientation = element.Orientation?.Value == S.OrientationValues.Landscape
                    ? ExcelPageOrientation.Landscape
                    : ExcelPageOrientation.Portrait,
                PaperSize = (ExcelPaperSize)(element.PaperSize?.Value ?? (uint)ExcelPaperSize.A4),
                Scale = element.Scale?.Value is uint scale ? (int)scale : null,
                FitToWidth = element.FitToWidth?.Value is uint width ? (int)width : null,
                FitToHeight = element.FitToHeight?.Value is uint height ? (int)height : null,
            };
        }
        set
        {
            _workbook.ThrowIfDisposed();
            S.Worksheet worksheet = Part().Worksheet!;
            S.PageSetup? element = worksheet.GetFirstChild<S.PageSetup>();
            if (value is null)
            {
                element?.Remove();
                return;
            }

            if (element is null)
            {
                element = new S.PageSetup();
                ExcelWorksheetLayout.InsertInOrder(worksheet, element);
            }

            element.PaperSize = (uint)value.PaperSize;
            element.Orientation = value.Orientation == ExcelPageOrientation.Landscape
                ? S.OrientationValues.Landscape
                : S.OrientationValues.Portrait;
            element.Scale = value.Scale is int scale ? (uint)scale : null;
            element.FitToWidth = value.FitToWidth is int fitWidth ? (uint)fitWidth : null;
            element.FitToHeight = value.FitToHeight is int fitHeight ? (uint)fitHeight : null;
            SetFitToPage(worksheet, value.FitToWidth is not null || value.FitToHeight is not null);
        }
    }

    /// <summary>Gets or sets the print options of the sheet.</summary>
    /// <value>
    /// The options, or <see langword="null"/> when the sheet has none. Assigning
    /// <see langword="null"/> removes them.
    /// </value>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelPrintOptions? PrintOptions
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.PrintOptions? element = Part().Worksheet!.GetFirstChild<S.PrintOptions>();
            if (element is null)
            {
                return null;
            }

            return new ExcelPrintOptions(
                element.HorizontalCentered?.Value ?? false,
                element.VerticalCentered?.Value ?? false,
                element.GridLines?.Value ?? false,
                element.Headings?.Value ?? false);
        }
        set
        {
            _workbook.ThrowIfDisposed();
            S.Worksheet worksheet = Part().Worksheet!;
            S.PrintOptions? element = worksheet.GetFirstChild<S.PrintOptions>();
            if (value is null)
            {
                element?.Remove();
                return;
            }

            if (element is null)
            {
                element = new S.PrintOptions();
                ExcelWorksheetLayout.InsertInOrder(worksheet, element);
            }

            element.HorizontalCentered = value.HorizontalCentered;
            element.VerticalCentered = value.VerticalCentered;
            element.GridLines = value.GridLines;
            element.Headings = value.Headings;
        }
    }

    /// <summary>Gets the hyperlinks of the worksheet.</summary>
    /// <exception cref="InvalidOfficeFileException">A stored hyperlink is malformed.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public IReadOnlyList<ExcelHyperlink> Hyperlinks
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.Hyperlinks? hyperlinks = Part().Worksheet!.GetFirstChild<S.Hyperlinks>();
            if (hyperlinks is null)
            {
                return [];
            }

            List<ExcelHyperlink> result = [];
            foreach (S.Hyperlink element in hyperlinks.Elements<S.Hyperlink>())
            {
                string reference = element.Reference?.Value ?? string.Empty;
                string? target = null;
                if (element.Id?.Value is string relationshipId)
                {
                    target = Part().HyperlinkRelationships
                        .FirstOrDefault(relationship => relationship.Id == relationshipId)
                        ?.Uri
                        .ToString();
                }

                result.Add(new ExcelHyperlink(
                    reference,
                    target,
                    element.Location?.Value,
                    element.Display?.Value,
                    element.Tooltip?.Value));
            }

            return result;
        }
    }

    /// <summary>Adds a hyperlink to an address outside the workbook.</summary>
    /// <param name="reference">The A1-style range the hyperlink covers, for example <c>A1</c>.</param>
    /// <param name="target">The address the hyperlink points to, for example <c>https://example.com</c>.</param>
    /// <param name="display">The text shown for the hyperlink.</param>
    /// <param name="tooltip">The tooltip of the hyperlink.</param>
    /// <returns>The hyperlink that was added.</returns>
    /// <exception cref="ArgumentException">
    /// The range is not a valid A1-style reference or the address is empty.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelHyperlink AddHyperlink(
        string reference,
        string target,
        string? display = null,
        string? tooltip = null)
    {
        string normalized = CellReference.NormalizeRangeOrCell(reference, nameof(reference));
        ArgumentException.ThrowIfNullOrEmpty(target);
        if (!Uri.TryCreate(target, UriKind.Absolute, out Uri? address))
        {
            throw new ArgumentException(
                "The address of an external hyperlink must be an absolute address.",
                nameof(target));
        }

        _workbook.ThrowIfDisposed();
        WorksheetPart part = Part();
        HyperlinkRelationship relationship = part.AddHyperlinkRelationship(address, true);

        S.Hyperlink element = new()
        {
            Reference = normalized,
            Id = relationship.Id,
        };
        if (display is not null)
        {
            element.Display = display;
        }

        if (tooltip is not null)
        {
            element.Tooltip = tooltip;
        }

        HyperlinksContainer(part.Worksheet!).AppendChild(element);
        return new ExcelHyperlink(normalized, address.ToString(), null, display, tooltip);
    }

    /// <summary>Adds a hyperlink to a place inside the workbook.</summary>
    /// <param name="reference">The A1-style range the hyperlink covers, for example <c>A1</c>.</param>
    /// <param name="location">
    /// The place the hyperlink points to, for example <c>Resumen!A1</c> or <c>A1</c>.
    /// </param>
    /// <param name="display">The text shown for the hyperlink.</param>
    /// <param name="tooltip">The tooltip of the hyperlink.</param>
    /// <returns>The hyperlink that was added.</returns>
    /// <exception cref="ArgumentException">
    /// The range is not a valid A1-style reference or the location is empty.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelHyperlink AddInternalHyperlink(
        string reference,
        string location,
        string? display = null,
        string? tooltip = null)
    {
        string normalized = CellReference.NormalizeRangeOrCell(reference, nameof(reference));
        ArgumentException.ThrowIfNullOrEmpty(location);
        _workbook.ThrowIfDisposed();

        S.Hyperlink element = new()
        {
            Reference = normalized,
            Location = location,
        };
        if (display is not null)
        {
            element.Display = display;
        }

        if (tooltip is not null)
        {
            element.Tooltip = tooltip;
        }

        HyperlinksContainer(Part().Worksheet!).AppendChild(element);
        return new ExcelHyperlink(normalized, null, location, display, tooltip);
    }

    /// <summary>Removes the hyperlink that covers a range.</summary>
    /// <param name="reference">The A1-style range of the hyperlink.</param>
    /// <returns><see langword="true"/> when a hyperlink was removed.</returns>
    /// <exception cref="ArgumentException"><paramref name="reference"/> is not a valid A1-style reference.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool RemoveHyperlink(string reference)
    {
        string normalized = CellReference.NormalizeRangeOrCell(reference, nameof(reference));
        _workbook.ThrowIfDisposed();

        WorksheetPart part = Part();
        S.Hyperlinks? hyperlinks = part.Worksheet!.GetFirstChild<S.Hyperlinks>();
        S.Hyperlink? target = hyperlinks?.Elements<S.Hyperlink>()
            .FirstOrDefault(element => string.Equals(element.Reference?.Value, normalized, StringComparison.OrdinalIgnoreCase));
        if (target is null)
        {
            return false;
        }

        if (target.Id?.Value is string relationshipId)
        {
            part.DeleteReferenceRelationship(relationshipId);
        }

        target.Remove();
        if (hyperlinks is not null && !hyperlinks.Elements<S.Hyperlink>().Any())
        {
            hyperlinks.Remove();
        }

        return true;
    }

    /// <summary>Gets the notes of the worksheet.</summary>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public IReadOnlyList<ExcelComment> Comments
    {
        get
        {
            _workbook.ThrowIfDisposed();
            List<ExcelComment> result = [];
            foreach ((string reference, string text, string author) in ReadCommentData())
            {
                result.Add(new ExcelComment(reference, text, author));
            }

            return result;
        }
    }

    /// <summary>Adds a note to a cell, replacing the one it had.</summary>
    /// <param name="reference">The A1-style reference of the cell, for example <c>A1</c>.</param>
    /// <param name="text">The text of the note.</param>
    /// <param name="author">The author of the note.</param>
    /// <returns>The note that was added.</returns>
    /// <exception cref="ArgumentException">
    /// The reference is not a valid A1-style reference, or the text or the author are empty.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelComment AddComment(string reference, string text, string author = "OfficeEngine")
    {
        string normalized = CellReference.NormalizeRangeOrCell(reference, nameof(reference));
        ArgumentException.ThrowIfNullOrEmpty(text);
        ArgumentException.ThrowIfNullOrEmpty(author);
        _workbook.ThrowIfDisposed();

        List<(string Reference, string Text, string Author)> comments = ReadCommentData();
        comments.RemoveAll(comment =>
            string.Equals(comment.Reference, normalized, StringComparison.OrdinalIgnoreCase));
        comments.Add((normalized, text, author));
        WriteComments(comments);
        return new ExcelComment(normalized, text, author);
    }

    /// <summary>Removes the note of a cell.</summary>
    /// <param name="reference">The A1-style reference of the cell.</param>
    /// <returns><see langword="true"/> when a note was removed.</returns>
    /// <exception cref="ArgumentException"><paramref name="reference"/> is not a valid A1-style reference.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool RemoveComment(string reference)
    {
        string normalized = CellReference.NormalizeRangeOrCell(reference, nameof(reference));
        _workbook.ThrowIfDisposed();

        List<(string Reference, string Text, string Author)> comments = ReadCommentData();
        int removed = comments.RemoveAll(comment =>
            string.Equals(comment.Reference, normalized, StringComparison.OrdinalIgnoreCase));
        if (removed == 0)
        {
            return false;
        }

        WriteComments(comments);
        return true;
    }

    /// <summary>Gets the images anchored in the worksheet.</summary>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public IReadOnlyList<ExcelImage> Images
    {
        get
        {
            _workbook.ThrowIfDisposed();
            DrawingsPart? drawing = Part().DrawingsPart;
            return drawing is null ? [] : ExcelImages.Read(drawing);
        }
    }

    /// <summary>Anchors an image in a cell.</summary>
    /// <param name="cellReference">The A1-style reference of the cell the image is anchored to.</param>
    /// <param name="content">The bytes of the image.</param>
    /// <param name="format">The format of the image.</param>
    /// <param name="widthPoints">The width of the image in points, greater than 0.</param>
    /// <param name="heightPoints">The height of the image in points, greater than 0.</param>
    /// <param name="name">The name of the image. A name is generated when it is not given.</param>
    /// <returns>The image that was added.</returns>
    /// <exception cref="ArgumentException">
    /// The reference is not a valid A1-style reference, the content is empty or a size is not
    /// positive.
    /// </exception>
    /// <exception cref="ArgumentNullException"><paramref name="content"/> is <see langword="null"/>.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelImage AddImage(
        string cellReference,
        Stream content,
        ExcelImageFormat format,
        double widthPoints,
        double heightPoints,
        string? name = null)
    {
        ArgumentNullException.ThrowIfNull(content);
        if (!CellReference.TryParse(cellReference, out int column, out int row))
        {
            throw new ArgumentException(
                $"'{cellReference}' is not a valid A1-style cell reference.",
                nameof(cellReference));
        }

        if (!double.IsFinite(widthPoints) || widthPoints <= 0d || !double.IsFinite(heightPoints) || heightPoints <= 0d)
        {
            throw new ArgumentException("The size of the image must be greater than zero.", nameof(widthPoints));
        }

        if (content.Length == 0)
        {
            throw new ArgumentException("The image is empty.", nameof(content));
        }

        _workbook.ThrowIfDisposed();
        WorksheetPart part = Part();
        DrawingsPart drawingPart = part.DrawingsPart ?? part.AddNewPart<DrawingsPart>();
        if (drawingPart.WorksheetDrawing is null)
        {
            drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            part.Worksheet!.AppendChild(new S.Drawing { Id = part.GetIdOfPart(drawingPart) });
        }

        ImagePart imagePart = drawingPart.AddImagePart(ToImagePartType(format));
        using (Stream target = imagePart.GetStream(FileMode.Create, FileAccess.Write))
        {
            content.CopyTo(target);
        }

        string imageName = name ?? $"Imagen {drawingPart.WorksheetDrawing.ChildElements.Count + 1}";
        uint id = (uint)(drawingPart.WorksheetDrawing.ChildElements.Count + 2);
        long width = (long)Math.Round(widthPoints * ExcelImages.EmusPerPoint);
        long height = (long)Math.Round(heightPoints * ExcelImages.EmusPerPoint);
        drawingPart.WorksheetDrawing.AppendChild(ExcelImages.BuildAnchor(
            id,
            imageName,
            drawingPart.GetIdOfPart(imagePart),
            column - 1,
            row - 1,
            width,
            height));

        return new ExcelImage(imageName, CellReference.ToReference(column, row), widthPoints, heightPoints, MediaTypeOf(format));
    }

    /// <summary>Removes an image of the worksheet by its name.</summary>
    /// <param name="name">The name of the image.</param>
    /// <returns><see langword="true"/> when an image was removed.</returns>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool RemoveImage(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        _workbook.ThrowIfDisposed();

        WorksheetPart part = Part();
        DrawingsPart? drawingPart = part.DrawingsPart;
        if (drawingPart?.WorksheetDrawing is null)
        {
            return false;
        }

        foreach (Xdr.OneCellAnchor anchor in drawingPart.WorksheetDrawing.Elements<Xdr.OneCellAnchor>().ToList())
        {
            Xdr.Picture? picture = anchor.GetFirstChild<Xdr.Picture>();
            string imageName = picture?.NonVisualPictureProperties?
                .NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            if (!string.Equals(imageName, name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (picture?.BlipFill?.Blip?.Embed?.Value is string relationshipId
                && drawingPart.GetPartById(relationshipId) is ImagePart imagePart)
            {
                drawingPart.DeletePart(imagePart);
            }

            anchor.Remove();
            return true;
        }

        return false;
    }

    /// <summary>Gets the charts anchored in the worksheet.</summary>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public IReadOnlyList<ExcelChart> Charts
    {
        get
        {
            _workbook.ThrowIfDisposed();
            DrawingsPart? drawing = Part().DrawingsPart;
            return drawing is null ? [] : ExcelCharts.Read(drawing);
        }
    }

    /// <summary>Adds a chart with a single series to the worksheet.</summary>
    /// <param name="cellReference">The A1-style reference of the cell the chart is anchored to.</param>
    /// <param name="type">The kind of chart.</param>
    /// <param name="categoriesRange">The A1-style range that holds the categories.</param>
    /// <param name="valuesRange">The A1-style range that holds the values.</param>
    /// <param name="title">The title of the chart.</param>
    /// <param name="seriesName">The name of the series.</param>
    /// <param name="widthPoints">The width of the chart in points.</param>
    /// <param name="heightPoints">The height of the chart in points.</param>
    /// <returns>The chart that was added.</returns>
    /// <remarks>
    /// The chart keeps a copy of the categories and the values as they are at the moment of the call,
    /// and points at the ranges, so that the engine that opens the file can recalculate it.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// A reference is not a valid A1-style reference or a size is not positive.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public ExcelChart AddChart(
        string cellReference,
        ExcelChartType type,
        string categoriesRange,
        string valuesRange,
        string? title = null,
        string? seriesName = null,
        double widthPoints = 480d,
        double heightPoints = 288d)
    {
        if (!CellReference.TryParse(cellReference, out int column, out int row))
        {
            throw new ArgumentException(
                $"'{cellReference}' is not a valid A1-style cell reference.",
                nameof(cellReference));
        }

        ArgumentNullException.ThrowIfNull(categoriesRange);
        ArgumentNullException.ThrowIfNull(valuesRange);

        if (!CellReference.TryParseRange(categoriesRange, out int firstCategoryColumn, out int firstCategoryRow, out _, out _)
            || !CellReference.TryParseRange(valuesRange, out int firstValueColumn, out int firstValueRow, out _, out _))
        {
            throw new ArgumentException(
                "The categories and the values must be A1-style ranges.",
                nameof(categoriesRange));
        }

        if (!double.IsFinite(widthPoints) || widthPoints <= 0d || !double.IsFinite(heightPoints) || heightPoints <= 0d)
        {
            throw new ArgumentException("The size of the chart must be greater than zero.", nameof(widthPoints));
        }

        _workbook.ThrowIfDisposed();
        WorksheetPart part = Part();
        DrawingsPart drawingPart = part.DrawingsPart ?? part.AddNewPart<DrawingsPart>();
        if (drawingPart.WorksheetDrawing is null)
        {
            drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            part.Worksheet!.AppendChild(new S.Drawing { Id = part.GetIdOfPart(drawingPart) });
        }

        ChartPart chartPart = drawingPart.AddNewPart<ChartPart>();
        string chartName = $"Gráfico {drawingPart.WorksheetDrawing.ChildElements.Count + 1}";
        uint id = (uint)(drawingPart.WorksheetDrawing.ChildElements.Count + 2);

        ExcelCharts.Build(
            chartPart,
            type,
            AbsoluteFormula(Name, firstCategoryColumn, firstCategoryRow, LastRowOf(categoriesRange), LastColumnOf(categoriesRange)),
            AbsoluteFormula(Name, firstValueColumn, firstValueRow, LastRowOf(valuesRange), LastColumnOf(valuesRange)),
            ReadRangeTexts(categoriesRange),
            ReadRangeNumbers(valuesRange),
            seriesName,
            title);

        long width = (long)Math.Round(widthPoints * ExcelImages.EmusPerPoint);
        long height = (long)Math.Round(heightPoints * ExcelImages.EmusPerPoint);
        Xdr.OneCellAnchor anchor = new(
            new Xdr.FromMarker(
                new Xdr.ColumnId((column - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((row - 1).ToString(CultureInfo.InvariantCulture)),
                new Xdr.RowOffset("0")),
            new Xdr.Extent { Cx = width, Cy = height },
            ExcelCharts.BuildFrame(id, chartName, drawingPart.GetIdOfPart(chartPart), width, height),
            new Xdr.ClientData());
        drawingPart.WorksheetDrawing.AppendChild(anchor);

        return new ExcelChart(
            chartName,
            CellReference.ToReference(column, row),
            type,
            categoriesRange,
            valuesRange,
            title,
            seriesName);
    }

    /// <summary>Removes a chart of the worksheet by its name.</summary>
    /// <param name="name">The name of the chart.</param>
    /// <returns><see langword="true"/> when a chart was removed.</returns>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public bool RemoveChart(string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);
        _workbook.ThrowIfDisposed();

        WorksheetPart part = Part();
        DrawingsPart? drawingPart = part.DrawingsPart;
        if (drawingPart?.WorksheetDrawing is null)
        {
            return false;
        }

        foreach (Xdr.OneCellAnchor anchor in drawingPart.WorksheetDrawing.Elements<Xdr.OneCellAnchor>().ToList())
        {
            Xdr.GraphicFrame? frame = anchor.GetFirstChild<Xdr.GraphicFrame>();
            string chartName = frame?.NonVisualGraphicFrameProperties?
                .NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            if (!string.Equals(chartName, name, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (frame?.Graphic?.GraphicData?.GetFirstChild<C.ChartReference>()?.Id?.Value is string relationshipId
                && drawingPart.GetPartById(relationshipId) is ChartPart chartPart)
            {
                drawingPart.DeletePart(chartPart);
            }

            anchor.Remove();
            return true;
        }

        return false;
    }

    /// <inheritdoc />
    public override string ToString() => Name;

    internal WorksheetPart Part()
    {
        string? relationshipId = _sheet.Id?.Value
            ?? throw new InvalidOfficeFileException($"Worksheet '{Name}' has no relationship identifier.");
        if (_workbook.WorkbookPart.GetPartById(relationshipId) is not WorksheetPart part)
        {
            throw new InvalidOfficeFileException($"Worksheet '{Name}' has no readable worksheet part.");
        }

        _ = OoxmlPartReader.Read($"the worksheet '{Name}'", () => part.Worksheet);
        return part;
    }

    internal S.SheetData SheetData()
    {
        WorksheetPart part = Part();
        S.Worksheet root = OoxmlPartReader.Read($"the worksheet '{Name}'", () => part.Worksheet);
        return root.GetFirstChild<S.SheetData>()
            ?? root.AppendChild(new S.SheetData());
    }

    internal S.Cell? FindCell(int column, int row)
    {
        string reference = CellReference.ToReference(column, row);
        S.SheetData sheetData = SheetData();
        S.Row? rowElement = FindRow(sheetData, row);
        if (rowElement is null)
        {
            EnsureNoRowsWithoutIndex(sheetData);
            return null;
        }

        return rowElement.Elements<S.Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, reference, StringComparison.Ordinal));
    }

    internal S.Cell GetOrCreateCell(int column, int row)
    {
        S.SheetData sheetData = SheetData();
        S.Row rowElement = FindRow(sheetData, row) ?? InsertRow(sheetData, row);

        string reference = CellReference.ToReference(column, row);
        S.Cell? cell = rowElement.Elements<S.Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, reference, StringComparison.Ordinal));
        if (cell is not null)
        {
            return cell;
        }

        cell = new S.Cell { CellReference = reference };
        S.Cell? next = rowElement.Elements<S.Cell>().FirstOrDefault(c => ColumnOf(c) > column);
        if (next is null)
        {
            rowElement.AppendChild(cell);
        }
        else
        {
            rowElement.InsertBefore(cell, next);
        }

        return cell;
    }

    private static int ColumnOf(S.Cell cell)
        => cell.CellReference?.Value is string reference && CellReference.TryParse(reference, out int column, out _) ? column : 0;

    private static PartTypeInfo ToImagePartType(ExcelImageFormat format) => format switch
    {
        ExcelImageFormat.Png => ImagePartType.Png,
        ExcelImageFormat.Jpeg => ImagePartType.Jpeg,
        ExcelImageFormat.Gif => ImagePartType.Gif,
        ExcelImageFormat.Bmp => ImagePartType.Bmp,
        ExcelImageFormat.Tiff => ImagePartType.Tiff,
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, "The image format is not defined."),
    };

    private static string MediaTypeOf(ExcelImageFormat format) => format switch
    {
        ExcelImageFormat.Png => "image/png",
        ExcelImageFormat.Jpeg => "image/jpeg",
        ExcelImageFormat.Gif => "image/gif",
        ExcelImageFormat.Bmp => "image/bmp",
        ExcelImageFormat.Tiff => "image/tiff",
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, "The image format is not defined."),
    };

    private static int LastRowOf(string range)
    {
        if (!CellReference.TryParseRange(range, out _, out _, out _, out int lastRow))
        {
            throw new ArgumentException($"'{range}' is not a valid A1-style range reference.", nameof(range));
        }

        return lastRow;
    }

    private static int LastColumnOf(string range)
    {
        if (!CellReference.TryParseRange(range, out _, out _, out int lastColumn, out _))
        {
            throw new ArgumentException($"'{range}' is not a valid A1-style range reference.", nameof(range));
        }

        return lastColumn;
    }

    private static string AbsoluteFormula(string sheetName, int firstColumn, int firstRow, int lastColumn, int lastRow)
        => ExcelDefinedNames.QuoteSheetName(sheetName) + "!" +
            ExcelDefinedNames.FormatAbsolute(firstColumn, firstRow, lastColumn, lastRow);

    private List<string?> ReadRangeTexts(string range)
    {
        List<string?> result = [];
        if (!CellReference.TryParseRange(range, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow))
        {
            throw new ArgumentException($"'{range}' is not a valid A1-style range reference.", nameof(range));
        }

        for (int row = firstRow; row <= lastRow; row++)
        {
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                result.Add(ReadCellText(column, row));
            }
        }

        return result;
    }

    private List<double?> ReadRangeNumbers(string range)
    {
        List<double?> result = [];
        if (!CellReference.TryParseRange(range, out int firstColumn, out int firstRow, out int lastColumn, out int lastRow))
        {
            throw new ArgumentException($"'{range}' is not a valid A1-style range reference.", nameof(range));
        }

        for (int row = firstRow; row <= lastRow; row++)
        {
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                object? value = GetCell(row, column).Value;
                result.Add(value is double number ? number : null);
            }
        }

        return result;
    }

    private List<(string Reference, string Text, string Author)> ReadCommentData()
    {
        WorksheetCommentsPart? commentsPart = Part().WorksheetCommentsPart;
        if (commentsPart?.Comments is null)
        {
            return [];
        }

        return ExcelComments.Read(commentsPart.Comments);
    }

    private void WriteComments(List<(string Reference, string Text, string Author)> comments)
    {
        WorksheetPart part = Part();
        WorksheetCommentsPart? commentsPart = part.WorksheetCommentsPart;
        if (commentsPart is null)
        {
            commentsPart = part.AddNewPart<WorksheetCommentsPart>();
        }

        commentsPart.Comments = ExcelComments.Build(comments, out _);

        VmlDrawingPart? drawing = part.VmlDrawingParts.FirstOrDefault();
        if (drawing is null)
        {
            drawing = part.AddNewPart<VmlDrawingPart>();
            S.Worksheet worksheet = part.Worksheet!;
            ExcelWorksheetLayout.InsertInOrder(worksheet, new S.LegacyDrawing
            {
                Id = part.GetIdOfPart(drawing),
            });
        }

        using Stream stream = drawing.GetStream(FileMode.Create, FileAccess.Write);
        using StreamWriter writer = new(stream, new UTF8Encoding(false));
        writer.Write(ExcelComments.BuildDrawing(comments));
    }

    private static S.Hyperlinks HyperlinksContainer(S.Worksheet worksheet)
    {
        S.Hyperlinks? hyperlinks = worksheet.GetFirstChild<S.Hyperlinks>();
        if (hyperlinks is null)
        {
            hyperlinks = new S.Hyperlinks();
            ExcelWorksheetLayout.InsertInOrder(worksheet, hyperlinks);
        }

        return hyperlinks;
    }

    private static void SetFitToPage(S.Worksheet worksheet, bool enabled)
    {
        S.SheetProperties? properties = worksheet.GetFirstChild<S.SheetProperties>();
        if (properties is null)
        {
            if (!enabled)
            {
                return;
            }

            properties = new S.SheetProperties();
            worksheet.InsertAt(properties, 0);
        }

        S.PageSetupProperties? pageSetup = properties.GetFirstChild<S.PageSetupProperties>();
        if (pageSetup is null)
        {
            if (!enabled)
            {
                return;
            }

            pageSetup = new S.PageSetupProperties();
            properties.AppendChild(pageSetup);
        }

        pageSetup.FitToPage = enabled;
    }

    private S.SheetView? FindSheetView()
    {
        S.SheetViews? views = Part().Worksheet!.GetFirstChild<S.SheetViews>();
        return views?.Elements<S.SheetView>().FirstOrDefault();
    }

    private static S.SheetView GetOrCreateSheetView(S.Worksheet worksheet)
    {
        S.SheetViews? views = worksheet.GetFirstChild<S.SheetViews>();
        if (views is null)
        {
            views = new S.SheetViews();
            ExcelWorksheetLayout.InsertInOrder(worksheet, views);
        }

        S.SheetView? view = views.Elements<S.SheetView>().FirstOrDefault();
        if (view is null)
        {
            view = new S.SheetView { WorkbookViewId = 0U };
            views.AppendChild(view);
        }

        return view;
    }

    private static S.Columns GetOrCreateColumns(S.Worksheet worksheet)
    {
        S.Columns? columns = worksheet.GetFirstChild<S.Columns>();
        if (columns is null)
        {
            columns = new S.Columns();
            ExcelWorksheetLayout.InsertInOrder(worksheet, columns);
        }

        return columns;
    }

    /// <summary>
    /// Returns the column element of a single column, splitting the one that covers it when the file
    /// groups several columns in a single element.
    /// </summary>
    private static S.Column EnsureColumn(S.Columns columns, int column)
    {
        S.Column? existing = columns.Elements<S.Column>()
            .FirstOrDefault(candidate => candidate.Min?.Value <= column && candidate.Max?.Value >= column);
        if (existing is null)
        {
            S.Column created = new() { Min = (uint)column, Max = (uint)column };
            InsertColumnInOrder(columns, created);
            return created;
        }

        uint min = existing.Min?.Value ?? (uint)column;
        uint max = existing.Max?.Value ?? (uint)column;
        if (min == column && max == column)
        {
            return existing;
        }

        S.Column middle = (S.Column)existing.CloneNode(true);
        middle.Min = (uint)column;
        middle.Max = (uint)column;
        existing.Remove();

        if (min < column)
        {
            S.Column before = (S.Column)middle.CloneNode(true);
            before.Min = min;
            before.Max = (uint)column - 1;
            InsertColumnInOrder(columns, before);
        }

        if (column < max)
        {
            S.Column after = (S.Column)middle.CloneNode(true);
            after.Min = (uint)column + 1;
            after.Max = max;
            InsertColumnInOrder(columns, after);
        }

        InsertColumnInOrder(columns, middle);
        return middle;
    }

    private static void InsertColumnInOrder(S.Columns columns, S.Column column)
    {
        S.Column? next = columns.Elements<S.Column>()
            .FirstOrDefault(candidate => candidate.Min?.Value > column.Min?.Value);
        if (next is null)
        {
            columns.AppendChild(column);
        }
        else
        {
            columns.InsertBefore(column, next);
        }
    }

    private static S.DefinedName? FindDefinedNameElement(S.DefinedNames names, string name, uint? localSheetId)        => names.Elements<S.DefinedName>().FirstOrDefault(element =>
            string.Equals(element.Name?.Value, name, StringComparison.OrdinalIgnoreCase)
            && element.LocalSheetId?.Value == localSheetId);

    private string? ReadCellText(int column, int row)
    {
        S.Cell? cell = FindCell(column, row);
        if (cell is null)
        {
            return null;
        }

        S.CellValues? dataType = cell.DataType?.Value;
        if (dataType == S.CellValues.InlineString)
        {
            return cell.InlineString?.GetFirstChild<S.Text>()?.Text;
        }

        string? raw = cell.CellValue?.Text;
        if (dataType == S.CellValues.SharedString)
        {
            if (raw is null
                || !int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                return null;
            }

            return _workbook.GetSharedString(index);
        }

        if (dataType == S.CellValues.Boolean)
        {
            return raw == "1" ? "TRUE" : "FALSE";
        }

        return raw;
    }

    private ExcelTable? FindTable(string name)
    {
        foreach (Worksheet sheet in _workbook.Worksheets)
        {
            foreach (ExcelTable table in sheet.Tables)
            {
                if (string.Equals(table.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return table;
                }
            }
        }

        return null;
    }

    private static string Unique(List<string> names, string candidate)
    {
        string result = candidate;
        int suffix = 1;
        while (names.Any(existing => string.Equals(existing, result, StringComparison.OrdinalIgnoreCase)))
        {
            suffix++;
            result = candidate + suffix.ToString(CultureInfo.InvariantCulture);
        }

        return result;
    }

    private static void ValidateIndex(int value, string parameterName, int max)
    {
        if (value < 1 || value > max)
        {
            throw new ArgumentOutOfRangeException(
                parameterName, value, $"The index must be between 1 and {max}.");
        }
    }

    private static S.Row? FindRow(S.SheetData sheetData, int row)
        => sheetData.Elements<S.Row>().FirstOrDefault(r => r.RowIndex is not null && r.RowIndex.Value == (uint)row);

    /// <summary>
    /// A row without the required <c>@r</c> attribute cannot be addressed reliably. Returning an
    /// empty cell here would silently hide the corruption, so the file is reported as unreadable.
    /// </summary>
    private static void EnsureNoRowsWithoutIndex(S.SheetData sheetData)
    {
        if (sheetData.Elements<S.Row>().Any(static r => r.RowIndex is null))
        {
            throw new InvalidOfficeFileException(
                "A worksheet row does not declare its required row index, so cell values cannot be read reliably.");
        }
    }

    private static S.Row InsertRow(S.SheetData sheetData, int row)
    {
        S.Row rowElement = new() { RowIndex = (uint)row };
        S.Row? next = sheetData.Elements<S.Row>().FirstOrDefault(r => r.RowIndex is not null && r.RowIndex.Value > (uint)row);
        if (next is null)
        {
            sheetData.AppendChild(rowElement);
        }
        else
        {
            sheetData.InsertBefore(rowElement, next);
        }

        return rowElement;
    }
}
