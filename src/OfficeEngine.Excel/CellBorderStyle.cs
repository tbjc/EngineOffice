namespace OfficeEngine.Excel;

/// <summary>
/// The line style of a cell border, mirroring the styles Office Open XML defines in
/// <c>ST_BorderStyle</c>.
/// </summary>
public enum CellBorderStyle
{
    /// <summary>No border.</summary>
    None,

    /// <summary>A thin line.</summary>
    Thin,

    /// <summary>A medium line.</summary>
    Medium,

    /// <summary>A dashed line.</summary>
    Dashed,

    /// <summary>A dotted line.</summary>
    Dotted,

    /// <summary>A thick line.</summary>
    Thick,

    /// <summary>A double line.</summary>
    DoubleLine,

    /// <summary>A hairline.</summary>
    Hair,

    /// <summary>A medium dashed line.</summary>
    MediumDashed,

    /// <summary>A dash-dot line.</summary>
    DashDot,

    /// <summary>A medium dash-dot line.</summary>
    MediumDashDot,

    /// <summary>A dash-dot-dot line.</summary>
    DashDotDot,

    /// <summary>A medium dash-dot-dot line.</summary>
    MediumDashDotDot,

    /// <summary>A slanted dash-dot line.</summary>
    SlantDashDot,
}
