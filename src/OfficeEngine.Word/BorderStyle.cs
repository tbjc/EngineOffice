namespace OfficeEngine.Word;

/// <summary>
/// The line style of a table or cell border, mirroring the styles Office Open XML defines in
/// <c>ST_Border</c>.
/// </summary>
/// <remarks>
/// <para>
/// The members cover the styles Word offers in its table border gallery. Office Open XML also defines
/// a large set of decorative "art" borders (<c>apples</c>, <c>balloons</c>, <c>hearts</c>…) that this
/// version does not model: reading one throws
/// <see cref="UnsupportedOfficeFeatureException"/> and the original markup is preserved.
/// </para>
/// <para>
/// The members named <c>SingleLine</c> and <c>DoubleLine</c> correspond to the Office Open XML values
/// <c>single</c> and <c>double</c>; the names avoid clashing with C# type names, exactly as
/// <see cref="UnderlineStyle"/> does.
/// </para>
/// </remarks>
public enum BorderStyle
{
    /// <summary>No border (<c>nil</c>): the edge is explicitly empty.</summary>
    Nil,

    /// <summary>No border (<c>none</c>).</summary>
    None,

    /// <summary>A single line (<c>single</c>).</summary>
    SingleLine,

    /// <summary>A thick line (<c>thick</c>).</summary>
    Thick,

    /// <summary>A double line (<c>double</c>).</summary>
    DoubleLine,

    /// <summary>A dotted line (<c>dotted</c>).</summary>
    Dotted,

    /// <summary>A dashed line (<c>dashed</c>).</summary>
    Dashed,

    /// <summary>A dot-dash line (<c>dotDash</c>).</summary>
    DotDash,

    /// <summary>A dot-dot-dash line (<c>dotDotDash</c>).</summary>
    DotDotDash,

    /// <summary>A triple line (<c>triple</c>).</summary>
    Triple,

    /// <summary>A thin outer line and a thick inner line, with a small gap (<c>thinThickSmallGap</c>).</summary>
    ThinThickSmallGap,

    /// <summary>A thick outer line and a thin inner line, with a small gap (<c>thickThinSmallGap</c>).</summary>
    ThickThinSmallGap,

    /// <summary>Thin, thick and thin lines, with small gaps (<c>thinThickThinSmallGap</c>).</summary>
    ThinThickThinSmallGap,

    /// <summary>A thin outer line and a thick inner line, with a medium gap (<c>thinThickMediumGap</c>).</summary>
    ThinThickMediumGap,

    /// <summary>A thick outer line and a thin inner line, with a medium gap (<c>thickThinMediumGap</c>).</summary>
    ThickThinMediumGap,

    /// <summary>Thin, thick and thin lines, with medium gaps (<c>thinThickThinMediumGap</c>).</summary>
    ThinThickThinMediumGap,

    /// <summary>A thin outer line and a thick inner line, with a large gap (<c>thinThickLargeGap</c>).</summary>
    ThinThickLargeGap,

    /// <summary>A thick outer line and a thin inner line, with a large gap (<c>thickThinLargeGap</c>).</summary>
    ThickThinLargeGap,

    /// <summary>Thin, thick and thin lines, with large gaps (<c>thinThickThinLargeGap</c>).</summary>
    ThinThickThinLargeGap,

    /// <summary>A wavy line (<c>wave</c>).</summary>
    Wave,

    /// <summary>A double wavy line (<c>doubleWave</c>).</summary>
    DoubleWave,

    /// <summary>A dashed line with a small gap (<c>dashSmallGap</c>).</summary>
    DashSmallGap,

    /// <summary>A dash-dot-stroked line (<c>dashDotStroked</c>).</summary>
    DashDotStroked,

    /// <summary>An embossed three-dimensional line (<c>threeDEmboss</c>).</summary>
    ThreeDEmboss,

    /// <summary>An engraved three-dimensional line (<c>threeDEngrave</c>).</summary>
    ThreeDEngrave,

    /// <summary>An outset three-dimensional line (<c>outset</c>).</summary>
    Outset,

    /// <summary>An inset three-dimensional line (<c>inset</c>).</summary>
    Inset,
}
