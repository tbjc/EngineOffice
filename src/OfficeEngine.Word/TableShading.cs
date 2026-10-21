using System;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// Shared access to the shading (<c>w:shd</c>) that a table and a cell hold.
/// </summary>
/// <remarks>
/// Only a solid fill is modelled, which Office Open XML expresses with the <c>clear</c> pattern.
/// A shading that uses another pattern fails with <see cref="UnsupportedOfficeFeatureException"/> and
/// leaves the original markup untouched; a shading whose fill is <c>auto</c>, or that carries a theme
/// fill instead of an explicit one, is reported as "no explicit colour".
/// </remarks>
internal static class TableShading
{
    /// <summary>Reads the fill colour a <c>w:shd</c> element stores.</summary>
    /// <param name="shading">The <c>w:shd</c> element, or <see langword="null"/>.</param>
    /// <returns>The stored colour, or <see langword="null"/> when there is none.</returns>
    /// <exception cref="InvalidOfficeFileException">A stored value cannot be parsed.</exception>
    /// <exception cref="UnsupportedOfficeFeatureException">
    /// The shading uses a pattern that Office Open XML defines but this version does not model.
    /// </exception>
    internal static RgbColor? Read(W.Shading? shading)
    {
        if (shading is null)
        {
            return null;
        }

        string? pattern = shading.Val?.InnerText;
        if (pattern is null)
        {
            throw new InvalidOfficeFileException("A table shading stores no w:val attribute, which this document cannot read.");
        }

        if (!string.Equals(pattern, "clear", StringComparison.Ordinal))
        {
            throw new UnsupportedOfficeFeatureException(
                $"A table shading uses the pattern '{pattern}', which this version does not model.");
        }

        string? fill = shading.Fill?.Value;
        if (fill is null || string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (!RgbColor.TryParseHex(fill, out RgbColor color))
        {
            throw new InvalidOfficeFileException($"A table shading stores the colour '{fill}' that this document cannot read.");
        }

        return color;
    }

    /// <summary>
    /// Writes a fill colour into a <c>w:shd</c> element, leaving any attribute this version does not
    /// model untouched.
    /// </summary>
    /// <param name="shading">The element to write.</param>
    /// <param name="color">The colour to store as the fill.</param>
    internal static void Write(W.Shading shading, RgbColor color)
    {
        shading.Val = W.ShadingPatternValues.Clear;
        shading.Fill = color.ToHex();
    }
}
