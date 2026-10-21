using System.Globalization;
using DocumentFormat.OpenXml;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>
/// A handle to a single cell of a <see cref="Worksheet"/>.
/// </summary>
/// <remarks>
/// <para>
/// The handle is a live view over the workbook. Reading a handle whose cell does not exist
/// yields <see langword="null"/> and does not create anything; assigning a non-null value creates
/// the row and cell on demand. Assigning <see langword="null"/> clears an existing cell and leaves
/// a missing one absent.
/// </para>
/// <para>
/// The supported value types are <see cref="string"/>, <see cref="bool"/>, <see cref="DateTime"/>,
/// <see langword="null"/> and numeric types that can be represented exactly as a <see cref="double"/>.
/// A date or time is stored as an Excel serial number and needs an explicit date/time
/// <see cref="NumberFormat"/>; use <see cref="SetDateTime"/> to set both in one call. A numeric cell
/// whose number format is a date/time format is read back as a <see cref="DateTime"/>.
/// </para>
/// <para>
/// Formulas can be written with <see cref="Formula"/> but are not calculated. A formula cell is read
/// as its cached value.
/// </para>
/// </remarks>
public sealed class Cell
{
    private readonly ExcelWorkbook _workbook;
    private readonly Worksheet _worksheet;
    private readonly int _column;
    private readonly int _row;

    internal Cell(ExcelWorkbook workbook, Worksheet worksheet, string reference, int column, int row)
    {
        _workbook = workbook;
        _worksheet = worksheet;
        Reference = reference;
        _column = column;
        _row = row;
    }

    /// <summary>Gets the A1-style reference of the cell, for example <c>B7</c>.</summary>
    public string Reference { get; }

    /// <summary>Gets the 1-based row index of the cell, for example <c>7</c> for <c>B7</c>.</summary>
    public int Row => _row;

    /// <summary>Gets the 1-based column index of the cell, for example <c>2</c> for <c>B7</c>.</summary>
    public int Column => _column;

    /// <summary>
    /// Gets or sets the value of the cell.
    /// </summary>
    /// <value>
    /// <see langword="null"/>, a <see cref="string"/>, a <see cref="bool"/>, a <see cref="DateTime"/>
    /// or a numeric value. Reading a stored value returns <see cref="string"/>, <see cref="double"/>,
    /// <see cref="DateTime"/>, <see cref="bool"/> or <see langword="null"/>. Formula cells are read as
    /// their cached value when present; use <see cref="Formula"/> for the formula text.
    /// </value>
    /// <remarks>
    /// Assigning a <see cref="DateTime"/> requires the cell to already have a date/time
    /// <see cref="NumberFormat"/>; otherwise use <see cref="SetDateTime"/> so the format is explicit.
    /// The date's wall-clock components are stored without converting its <see cref="DateTime.Kind"/>.
    /// Assigning any value replaces a formula in the cell; assigning <see langword="null"/> removes both
    /// the formula and its cached result.
    /// </remarks>
    /// <exception cref="InvalidOperationException">
    /// A <see cref="DateTime"/> is assigned without a date/time number format.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A <see cref="DateTime"/> is outside the workbook's date system.
    /// </exception>
    /// <exception cref="UnsupportedOfficeFeatureException">The assigned value has an unsupported type.</exception>
    /// <exception cref="InvalidOfficeFileException">The stored cell is malformed and cannot be read.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public object? Value
    {
        get => ReadValue();
        set => WriteValue(value);
    }

    /// <summary>
    /// Gets or sets the number format code of the cell, for example <c>0.00</c> or <c>yyyy-mm-dd</c>.
    /// </summary>
    /// <value>
    /// The format code, <c>General</c> when no explicit format applies, or <see langword="null"/> for a
    /// built-in format that this version does not model. Assigning <see langword="null"/> resets the
    /// cell to <c>General</c>.
    /// </value>
    /// <remarks>
    /// Setting a non-General format on a missing cell creates the cell so the format has somewhere to
    /// live; setting <c>General</c> or <see langword="null"/> on a missing cell leaves it absent.
    /// Existing font, fill, border and alignment properties of the cell are preserved.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// The assigned format is empty, longer than 255 characters or not valid XML.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">The stored style is malformed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The workbook has no room for another custom number format.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public string? NumberFormat
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.Cell? cell = _worksheet.FindCell(_column, _row);
            return cell is null ? ExcelStyles.General : ExcelStyles.Resolve(_workbook, cell).Code;
        }
        set
        {
            _workbook.ThrowIfDisposed();
            string normalized = ExcelStyles.Normalize(value);
            if (string.Equals(normalized, ExcelStyles.General, StringComparison.Ordinal))
            {
                S.Cell? existing = _worksheet.FindCell(_column, _row);
                if (existing is not null)
                {
                    ExcelStyles.Set(_workbook, existing, normalized);
                }

                return;
            }

            S.Cell target = _worksheet.GetOrCreateCell(_column, _row);
            ExcelStyles.Set(_workbook, target, normalized);
        }
    }

    /// <summary>
    /// Gets or sets the formula text of the cell.
    /// </summary>
    /// <value>
    /// The formula with a leading <c>=</c>, or <see langword="null"/> when the cell has no formula or
    /// its formula text is stored in another cell (a shared formula).
    /// </value>
    /// <remarks>
    /// <para>
    /// The setter accepts a formula with or without the leading <c>=</c> and stores a normal formula.
    /// It does not calculate or semantically validate the formula; this library never evaluates it.
    /// Assigning a formula removes any cached result, so <see cref="Value"/> is <see langword="null"/>
    /// until Excel or another engine recalculates the workbook. The workbook is marked for a full
    /// recalculation when it is opened.
    /// </para>
    /// <para>
    /// Assigning <see langword="null"/> removes the formula and its cached result from an existing cell;
    /// a missing cell stays absent. Setting a formula on a missing cell creates the cell. The number
    /// format and the cell style are preserved.
    /// </para>
    /// <para>
    /// Shared, array and data-table formulas are preserved when the workbook is read and saved, but this
    /// version writes only normal formulas. A shared formula whose text is stored in another cell reads
    /// as <see langword="null"/> because it is not expanded.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// The formula is empty, longer than 8192 characters or not valid XML.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">The stored workbook is malformed.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public string? Formula
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.Cell? cell = _worksheet.FindCell(_column, _row);
            return ExcelFormula.ToPublic(cell?.CellFormula?.Text);
        }
        set
        {
            _workbook.ThrowIfDisposed();
            if (value is null)
            {
                S.Cell? existing = _worksheet.FindCell(_column, _row);
                if (existing is not null)
                {
                    ClearContent(existing);
                }

                return;
            }

            string normalized = ExcelFormula.Normalize(value);
            S.Cell target = _worksheet.GetOrCreateCell(_column, _row);
            ClearContent(target);
            target.CellFormula = new S.CellFormula(normalized);
            ExcelFormula.RequestRecalculation(_workbook);
        }
    }

    /// <summary>
    /// Gets or sets the reusable font, fill, border and alignment formatting of the cell.
    /// </summary>
    /// <value>
    /// A <see cref="CellStyle"/> describing the modelled formatting. A missing cell reports an empty
    /// style. Assigning an empty style resets the modelled formatting to the workbook defaults.
    /// </value>
    /// <remarks>
    /// The style does not include the number format: <see cref="NumberFormat"/> is preserved when a
    /// style is applied. Setting a non-empty style on a missing cell creates the cell so the style
    /// has somewhere to live; setting an empty style on a missing cell leaves it absent. Applying
    /// the same style to several cells reuses one <c>cellXfs</c> entry when the rest of the cell
    /// format is the same.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// The font name is empty, too long or not valid XML.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The font size or an alignment value is outside the supported range.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">The stored style is malformed.</exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public CellStyle Style
    {
        get
        {
            _workbook.ThrowIfDisposed();
            S.Cell? cell = _worksheet.FindCell(_column, _row);
            return cell is null ? new CellStyle() : ExcelStyles.ResolveStyle(_workbook, cell);
        }
        set
        {
            _workbook.ThrowIfDisposed();
            ExcelStyles.ValidateStyle(value);
            S.Cell? existing = _worksheet.FindCell(_column, _row);
            if (existing is null)
            {
                if (ExcelStyles.IsDefaultStyle(value))
                {
                    return;
                }

                existing = _worksheet.GetOrCreateCell(_column, _row);
            }

            ExcelStyles.ApplyStyle(_workbook, existing, value);
        }
    }

    /// <summary>Writes a date/time value with an explicit date/time number format.</summary>
    /// <param name="value">
    /// The date and time to store. Its wall-clock components are stored without converting its
    /// <see cref="DateTime.Kind"/>.
    /// </param>
    /// <param name="format">A date/time format code such as <c>yyyy-mm-dd</c> or <c>h:mm</c>.</param>
    /// <remarks>
    /// The value is stored as an Excel serial number and the format is applied to the cell. The format
    /// must contain at least one date or time token.
    /// </remarks>
    /// <exception cref="ArgumentException">
    /// <paramref name="format"/> is empty, too long, not valid XML or not a date/time format.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="value"/> is outside the workbook's date system.
    /// </exception>
    /// <exception cref="InvalidOfficeFileException">The stored style is malformed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The workbook has no room for another custom number format.
    /// </exception>
    /// <exception cref="ObjectDisposedException">The workbook has been disposed.</exception>
    public void SetDateTime(DateTime value, string format)
    {
        _workbook.ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrEmpty(format);
        string normalized = ExcelStyles.Normalize(format);
        if (string.Equals(normalized, ExcelStyles.General, StringComparison.Ordinal) ||
            !ExcelStyles.IsDateTime(normalized))
        {
            throw new ArgumentException("The format must contain date or time tokens.", nameof(format));
        }

        double serial = ToSerial(value);
        S.Cell target = _worksheet.GetOrCreateCell(_column, _row);
        ExcelStyles.Set(_workbook, target, normalized);
        WriteSerial(serial, target);
    }

    /// <inheritdoc />
    public override string ToString() => $"{Reference} = {Value ?? "(null)"}";

    private object? ReadValue()
    {
        _workbook.ThrowIfDisposed();
        S.Cell? cell = _worksheet.FindCell(_column, _row);
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
            if (raw is null)
            {
                return null;
            }

            if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                throw new InvalidOfficeFileException($"Cell {Reference} has an invalid shared string index '{raw}'.");
            }

            return _workbook.GetSharedString(index)
                ?? throw new InvalidOfficeFileException($"Cell {Reference} refers to shared string {index}, which does not exist.");
        }

        if (dataType == S.CellValues.Boolean)
        {
            return raw == "1";
        }

        if (dataType == S.CellValues.String || dataType == S.CellValues.Error)
        {
            return raw;
        }

        if (string.IsNullOrEmpty(raw))
        {
            return null;
        }

        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
        {
            return raw;
        }

        if (ExcelStyles.Resolve(_workbook, cell).IsDateTime &&
            TryFromSerial(number, _workbook.Uses1904DateSystem, out DateTime dateTime))
        {
            return dateTime;
        }

        return number;
    }

    private void WriteValue(object? value)
    {
        _workbook.ThrowIfDisposed();

        if (value is null)
        {
            ClearValue();
            return;
        }

        switch (value)
        {
            case string text:
            {
                S.Cell target = _worksheet.GetOrCreateCell(_column, _row);
                ClearContent(target);
                target.DataType = S.CellValues.SharedString;
                target.CellValue = new S.CellValue(
                    _workbook.GetOrAddSharedString(text).ToString(CultureInfo.InvariantCulture));
                return;
            }

            case bool flag:
            {
                S.Cell target = _worksheet.GetOrCreateCell(_column, _row);
                ClearContent(target);
                target.DataType = S.CellValues.Boolean;
                target.CellValue = new S.CellValue(flag ? "1" : "0");
                return;
            }

            case DateTime dateTime:
            {
                if (!HasDateTimeNumberFormat())
                {
                    throw new InvalidOperationException(
                        "A DateTime value needs an explicit date/time number format. " +
                        "Set NumberFormat to a date/time format or use SetDateTime(value, format).");
                }

                double serial = ToSerial(dateTime);
                S.Cell target = _worksheet.GetOrCreateCell(_column, _row);
                WriteSerial(serial, target);
                return;
            }

            default:
            {
                double number = ToNumber(value);
                S.Cell target = _worksheet.GetOrCreateCell(_column, _row);
                ClearContent(target);
                target.CellValue = new S.CellValue(number.ToString(CultureInfo.InvariantCulture));
                return;
            }
        }
    }

    /// <summary>
    /// Validates that <see cref="CellRange.SetValues"/> can store <paramref name="value"/> without
    /// mutating the range. A DateTime is rejected because each cell needs its own explicit format.
    /// </summary>
    internal static void EnsureRangeValue(object? value)
    {
        if (value is DateTime)
        {
            throw new UnsupportedOfficeFeatureException(
                "CellRange.SetValues does not write DateTime values because each cell needs an " +
                "explicit number format. Use Cell.SetDateTime(value, format).");
        }

        if (value is null or string or bool)
        {
            return;
        }

        _ = ToNumber(value);
    }

    /// <summary>
    /// Removes the stored value of an existing cell. A cell that does not exist is left absent
    /// instead of being created empty.
    /// </summary>
    internal void ClearValue()
    {
        _workbook.ThrowIfDisposed();
        S.Cell? target = _worksheet.FindCell(_column, _row);
        if (target is null)
        {
            return;
        }

        ClearContent(target);
    }

    private bool HasDateTimeNumberFormat()
    {
        S.Cell? cell = _worksheet.FindCell(_column, _row);
        return cell is not null && ExcelStyles.Resolve(_workbook, cell).IsDateTime;
    }

    /// <summary>
    /// Removes the formula, cached value and inline string from an existing cell. A removed formula
    /// leaves dependants without a valid cache, so the workbook asks for a full recalculation on open.
    /// </summary>
    private void ClearContent(S.Cell target)
    {
        bool hadFormula = target.CellFormula is not null;
        if (target.DataType?.Value == S.CellValues.SharedString)
        {
            _workbook.SharedStringReferences = Math.Max(0, _workbook.SharedStringReferences - 1);
        }

        target.RemoveAllChildren<S.CellFormula>();
        target.RemoveAllChildren<S.CellValue>();
        target.RemoveAllChildren<S.InlineString>();
        target.DataType = null;
        if (hadFormula)
        {
            ExcelFormula.RequestRecalculation(_workbook);
        }
    }

    private void WriteSerial(double serial, S.Cell target)
    {
        ClearContent(target);
        target.CellValue = new S.CellValue(serial.ToString("R", CultureInfo.InvariantCulture));
    }

    private double ToSerial(DateTime value)
    {
        if (_workbook.Uses1904DateSystem)
        {
            if (value < new DateTime(1904, 1, 1))
            {
                throw new ArgumentOutOfRangeException(
                    nameof(value),
                    value,
                    "Excel's 1904 date system starts on 1904-01-01.");
            }

            return value.ToOADate() - 1462d;
        }

        if (value < new DateTime(1900, 1, 1))
        {
            throw new ArgumentOutOfRangeException(
                nameof(value),
                value,
                "Excel's 1900 date system starts on 1900-01-01.");
        }

        double oaDate = value.ToOADate();
        return oaDate < 61d ? oaDate - 1d : oaDate;
    }

    private static bool TryFromSerial(double serial, bool uses1904DateSystem, out DateTime value)
    {
        value = default;
        if (!double.IsFinite(serial))
        {
            return false;
        }

        double oaDate;
        if (uses1904DateSystem)
        {
            if (serial < 0d || serial > MaxSerial1904)
            {
                return false;
            }

            oaDate = serial + 1462d;
        }
        else
        {
            // Excel's 1900 system has a fictitious 1900-02-29 (serial 60) and no serial 0.
            if (serial < 1d || (serial >= 60d && serial < 61d) || serial > MaxSerial1900)
            {
                return false;
            }

            oaDate = serial < 60d ? serial + 1d : serial;
        }

        try
        {
            value = DateTime.FromOADate(oaDate);
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private static double ToNumber(object value) => value switch
    {
        double number => number,
        float number => number,
        decimal number => (double)number,
        int number => number,
        uint number => number,
        long number => number,
        ulong number => number,
        short number => number,
        ushort number => number,
        byte number => number,
        sbyte number => number,
        _ => throw new UnsupportedOfficeFeatureException(
            $"Cell values of type '{value.GetType().FullName}' are not supported. " +
            "Supported types are string, bool, DateTime (with an explicit date/time number format), " +
            "null and numeric types that fit in a double."),
    };

    private static readonly double MaxSerial1900 = DateTime.MaxValue.ToOADate();
    private static readonly double MaxSerial1904 = DateTime.MaxValue.ToOADate() - 1462d;
}
