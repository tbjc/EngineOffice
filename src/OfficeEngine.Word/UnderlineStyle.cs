namespace OfficeEngine.Word;

/// <summary>
/// The character underline styles modelled by OfficeEngine.
/// </summary>
/// <remarks>
/// The members mirror the Office Open XML <c>ST_Underline</c> value set, so any underline a
/// Word document can contain maps to exactly one member and a round-trip is lossless.
/// </remarks>
public enum UnderlineStyle
{
    /// <summary>No underline (<c>none</c>).</summary>
    None,

    /// <summary>A single line (<c>single</c>).</summary>
    SingleLine,

    /// <summary>Underlines only the words, leaving the spaces unmarked (<c>words</c>).</summary>
    Words,

    /// <summary>A double line (<c>double</c>).</summary>
    DoubleLine,

    /// <summary>A thick line (<c>thick</c>).</summary>
    Thick,

    /// <summary>A dotted line (<c>dotted</c>).</summary>
    Dotted,

    /// <summary>A heavy dotted line (<c>dottedHeavy</c>).</summary>
    DottedHeavy,

    /// <summary>A dashed line (<c>dash</c>).</summary>
    Dash,

    /// <summary>A heavy dashed line (<c>dashedHeavy</c>).</summary>
    DashedHeavy,

    /// <summary>A long dashed line (<c>dashLong</c>).</summary>
    DashLong,

    /// <summary>A heavy long dashed line (<c>dashLongHeavy</c>).</summary>
    DashLongHeavy,

    /// <summary>A dot-dash line (<c>dotDash</c>).</summary>
    DotDash,

    /// <summary>A heavy dot-dash line (<c>dashDotHeavy</c>).</summary>
    DashDotHeavy,

    /// <summary>A dot-dot-dash line (<c>dotDotDash</c>).</summary>
    DotDotDash,

    /// <summary>A heavy dot-dot-dash line (<c>dashDotDotHeavy</c>).</summary>
    DashDotDotHeavy,

    /// <summary>A wavy line (<c>wave</c>).</summary>
    Wave,

    /// <summary>A heavy wavy line (<c>wavyHeavy</c>).</summary>
    WavyHeavy,

    /// <summary>A double wavy line (<c>wavyDouble</c>).</summary>
    WavyDouble,
}
