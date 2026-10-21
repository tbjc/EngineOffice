namespace OfficeEngine.Word;

/// <summary>
/// Kind of measure a table or cell width uses, mirroring a subset of OOXML <c>ST_TblWidth</c>.
/// </summary>
public enum TableWidthType
{
    /// <summary>The layout algorithm decides the width (OOXML <c>auto</c>).</summary>
    Auto,

    /// <summary>The width is a number of points, stored in twentieths of a point (OOXML <c>dxa</c>).</summary>
    Points,

    /// <summary>The width is a percentage of the available width (OOXML <c>pct</c>).</summary>
    Percentage,
}
