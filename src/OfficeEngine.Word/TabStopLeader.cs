namespace OfficeEngine.Word;

/// <summary>
/// Leader character drawn between the previous tab stop and a tab stop, mirroring OOXML
/// <c>ST_TabTlc</c>.
/// </summary>
public enum TabStopLeader
{
    /// <summary>No leader character.</summary>
    None,

    /// <summary>A dotted leader, as in a table of contents.</summary>
    Dot,

    /// <summary>A hyphen leader.</summary>
    Hyphen,

    /// <summary>An underscore leader.</summary>
    Underscore,

    /// <summary>A heavy (thick) leader line.</summary>
    Heavy,

    /// <summary>A middle-dot leader.</summary>
    MiddleDot,
}
