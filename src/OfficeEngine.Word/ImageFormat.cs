namespace OfficeEngine.Word;

/// <summary>
/// Image formats that OfficeEngine can embed in a Word document.
/// </summary>
/// <remarks>
/// Office Open XML supports more image formats than this version models. The ones listed here cover
/// the common raster and metafile formats; SVG, JPEG 2000, PCX and icons are preserved when they
/// already exist in a document but cannot be added through this API.
/// </remarks>
public enum ImageFormat
{
    /// <summary>Portable Network Graphics (<c>.png</c>).</summary>
    Png,

    /// <summary>JPEG (<c>.jpg</c>, <c>.jpeg</c>).</summary>
    Jpeg,

    /// <summary>Graphics Interchange Format (<c>.gif</c>).</summary>
    Gif,

    /// <summary>Windows bitmap (<c>.bmp</c>).</summary>
    Bmp,

    /// <summary>Tagged Image File Format (<c>.tif</c>, <c>.tiff</c>).</summary>
    Tiff,

    /// <summary>Enhanced Metafile (<c>.emf</c>).</summary>
    Emf,

    /// <summary>Windows Metafile (<c>.wmf</c>).</summary>
    Wmf,
}
