using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>
/// A defined name of a workbook: a name that points at a range of a worksheet, or at something this
/// version does not model (a constant, a formula, a built-in name such as a print area).
/// </summary>
/// <remarks>
/// <para>
/// The catalogue of defined names is read from the workbook and is not a copy of the file: the
/// values reflect what the workbook contains when they are read.
/// </para>
/// <para>
/// A name can be scoped to the whole workbook or to a single worksheet. A scoped name is reported
/// with its <see cref="Scope"/>; a workbook-scoped name has no scope.
/// </para>
/// </remarks>
public sealed class ExcelDefinedName
{
    internal ExcelDefinedName(ExcelWorkbook workbook, S.DefinedName element, Worksheet? scope)
    {
        Scope = scope;
        Name = element.Name?.Value ?? string.Empty;
        RefersTo = element.Text;
        IsRange = ExcelDefinedNames.TryParseTarget(workbook, scope, element.Text, out string? range);
        Range = range;
    }

    /// <summary>Gets the name of the defined name, for example <c>Datos</c>.</summary>
    public string Name { get; }

    /// <summary>
    /// Gets the text that the name refers to, exactly as it is stored, for example
    /// <c>Sheet1!$A$1:$B$2</c>.
    /// </summary>
    public string RefersTo { get; }

    /// <summary>
    /// Gets the worksheet the name is scoped to, or <see langword="null"/> when the name is scoped to
    /// the whole workbook.
    /// </summary>
    public Worksheet? Scope { get; }

    /// <summary>
    /// Gets whether the name points at a cell or a contiguous range of a worksheet of this workbook.
    /// </summary>
    /// <remarks>
    /// A name that refers to a constant, to a formula, to several areas or to another workbook is not
    /// a range; it is preserved but not interpreted.
    /// </remarks>
    public bool IsRange { get; }

    /// <summary>
    /// Gets the A1-style range the name points at, for example <c>A1:B2</c>, or <see langword="null"/>
    /// when the name is not a range.
    /// </summary>
    public string? Range { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Name} = {RefersTo}";
}
