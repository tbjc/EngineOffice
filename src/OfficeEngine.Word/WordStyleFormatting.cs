namespace OfficeEngine.Word;

/// <summary>
/// The formatting a style definition contributes, read from a WordprocessingML <c>w:style</c>.
/// </summary>
/// <remarks>
/// <para>
/// Every member is <em>tri-state</em>: <see langword="null"/> means the style does not define that
/// property and it is therefore inherited, which is exactly what Office Open XML expresses by
/// omitting the corresponding element or attribute.
/// </para>
/// <para>
/// Character properties come from the style's <c>w:rPr</c> (family and size of the font, bold,
/// italics and explicit colour) and paragraph properties from its <c>w:pPr</c> (alignment,
/// indentation and spacing before and after). Properties that Office Open XML supports but this
/// version does not model (highlighting, strike-through, borders, line spacing, numbering…) are not
/// reported here.
/// </para>
/// <para>
/// <see cref="RgbColor"/> is used for the colour, so a theme colour or the <c>auto</c> keyword is
/// reported as <see langword="null"/> ("no fixed RGB colour"), exactly as <see cref="Run.Color"/>
/// does. See <c>docs/adr/0008-character-formatting-model.md</c>.
/// </para>
/// </remarks>
public readonly struct WordStyleFormatting : IEquatable<WordStyleFormatting>
{
    internal WordStyleFormatting(
        string? fontFamily,
        double? fontSizePoints,
        bool? bold,
        bool? italic,
        RgbColor? color,
        ParagraphAlignment? alignment,
        double? leftIndentPoints,
        double? rightIndentPoints,
        double? firstLineIndentPoints,
        double? hangingIndentPoints,
        double? spacingBeforePoints,
        double? spacingAfterPoints)
    {
        FontFamily = fontFamily;
        FontSizePoints = fontSizePoints;
        Bold = bold;
        Italic = italic;
        Color = color;
        Alignment = alignment;
        LeftIndentPoints = leftIndentPoints;
        RightIndentPoints = rightIndentPoints;
        FirstLineIndentPoints = firstLineIndentPoints;
        HangingIndentPoints = hangingIndentPoints;
        SpacingBeforePoints = spacingBeforePoints;
        SpacingAfterPoints = spacingAfterPoints;
    }

    /// <summary>Gets the font family (<c>w:rFonts</c>), or <see langword="null"/> when it is inherited.</summary>
    public string? FontFamily { get; }

    /// <summary>Gets the font size in points (<c>w:sz</c>), or <see langword="null"/> when it is inherited.</summary>
    /// <remarks>Office Open XML stores the size in half-points, so a stored size is a multiple of <c>0.5</c>.</remarks>
    public double? FontSizePoints { get; }

    /// <summary>
    /// Gets whether the style sets the text bold (<c>w:b</c>), or <see langword="null"/> when boldness
    /// is inherited. <see langword="false"/> explicitly turns it off.
    /// </summary>
    public bool? Bold { get; }

    /// <summary>
    /// Gets whether the style sets the text italic (<c>w:i</c>), or <see langword="null"/> when
    /// italics are inherited. <see langword="false"/> explicitly turns them off.
    /// </summary>
    public bool? Italic { get; }

    /// <summary>
    /// Gets the explicit RGB colour (<c>w:color</c>), or <see langword="null"/> when the style has no
    /// fixed RGB colour (it is unspecified, <c>auto</c> or a theme colour).
    /// </summary>
    public RgbColor? Color { get; }

    /// <summary>Gets the paragraph alignment (<c>w:jc</c>), or <see langword="null"/> when it is inherited.</summary>
    public ParagraphAlignment? Alignment { get; }

    /// <summary>Gets the left indent in points (<c>w:ind/@w:left</c>), or <see langword="null"/> when it is inherited.</summary>
    public double? LeftIndentPoints { get; }

    /// <summary>Gets the right indent in points (<c>w:ind/@w:right</c>), or <see langword="null"/> when it is inherited.</summary>
    public double? RightIndentPoints { get; }

    /// <summary>Gets the first-line indent in points (<c>w:ind/@w:firstLine</c>), or <see langword="null"/> when it is inherited.</summary>
    public double? FirstLineIndentPoints { get; }

    /// <summary>Gets the hanging indent in points (<c>w:ind/@w:hanging</c>), or <see langword="null"/> when it is inherited.</summary>
    public double? HangingIndentPoints { get; }

    /// <summary>Gets the spacing before the paragraph in points (<c>w:spacing/@w:before</c>), or <see langword="null"/> when it is inherited.</summary>
    public double? SpacingBeforePoints { get; }

    /// <summary>Gets the spacing after the paragraph in points (<c>w:spacing/@w:after</c>), or <see langword="null"/> when it is inherited.</summary>
    public double? SpacingAfterPoints { get; }

    /// <inheritdoc />
    public bool Equals(WordStyleFormatting other)
        => FontFamily == other.FontFamily
            && FontSizePoints == other.FontSizePoints
            && Bold == other.Bold
            && Italic == other.Italic
            && Color == other.Color
            && Alignment == other.Alignment
            && LeftIndentPoints == other.LeftIndentPoints
            && RightIndentPoints == other.RightIndentPoints
            && FirstLineIndentPoints == other.FirstLineIndentPoints
            && HangingIndentPoints == other.HangingIndentPoints
            && SpacingBeforePoints == other.SpacingBeforePoints
            && SpacingAfterPoints == other.SpacingAfterPoints;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is WordStyleFormatting other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        HashCode hash = new();
        hash.Add(FontFamily);
        hash.Add(FontSizePoints);
        hash.Add(Bold);
        hash.Add(Italic);
        hash.Add(Color);
        hash.Add(Alignment);
        hash.Add(LeftIndentPoints);
        hash.Add(RightIndentPoints);
        hash.Add(FirstLineIndentPoints);
        hash.Add(HangingIndentPoints);
        hash.Add(SpacingBeforePoints);
        hash.Add(SpacingAfterPoints);
        return hash.ToHashCode();
    }

    /// <summary>Determines whether two formattings are equal.</summary>
    /// <param name="left">The first formatting.</param>
    /// <param name="right">The second formatting.</param>
    /// <returns><see langword="true"/> if both are equal; otherwise, <see langword="false"/>.</returns>
    public static bool operator ==(WordStyleFormatting left, WordStyleFormatting right) => left.Equals(right);

    /// <summary>Determines whether two formattings are different.</summary>
    /// <param name="left">The first formatting.</param>
    /// <param name="right">The second formatting.</param>
    /// <returns><see langword="true"/> if both are different; otherwise, <see langword="false"/>.</returns>
    public static bool operator !=(WordStyleFormatting left, WordStyleFormatting right) => !left.Equals(right);
}
