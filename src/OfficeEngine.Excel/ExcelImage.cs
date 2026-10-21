namespace OfficeEngine.Excel;

/// <summary>An image anchored in a worksheet.</summary>
/// <remarks>
/// The image is anchored to a cell and keeps the size it was inserted with, in points. The size of
/// the cells around it is not taken into account: the image does not grow or shrink with the sheet.
/// </remarks>
public sealed class ExcelImage
{
    internal ExcelImage(
        string name,
        string reference,
        double widthPoints,
        double heightPoints,
        string mediaType)
    {
        Name = name;
        Reference = reference;
        WidthPoints = widthPoints;
        HeightPoints = heightPoints;
        MediaType = mediaType;
    }

    /// <summary>Gets the name of the image, for example <c>Imagen 1</c>.</summary>
    public string Name { get; }

    /// <summary>Gets the A1-style reference of the cell the image is anchored to.</summary>
    public string Reference { get; }

    /// <summary>Gets the width of the image in points.</summary>
    public double WidthPoints { get; }

    /// <summary>Gets the height of the image in points.</summary>
    public double HeightPoints { get; }

    /// <summary>Gets the media type of the image, for example <c>image/png</c>.</summary>
    public string MediaType { get; }

    /// <inheritdoc />
    public override string ToString() => $"{Name} en {Reference}";
}
