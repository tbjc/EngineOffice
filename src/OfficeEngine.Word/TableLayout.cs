namespace OfficeEngine.Word;

/// <summary>
/// How a table lays its columns out, mirroring OOXML <c>ST_TblLayoutType</c>.
/// </summary>
public enum TableLayout
{
    /// <summary>Columns are resized to fit their content (OOXML <c>autofit</c>).</summary>
    AutoFit,

    /// <summary>Columns keep the widths the table grid declares (OOXML <c>fixed</c>).</summary>
    Fixed,
}
