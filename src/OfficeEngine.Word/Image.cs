using OfficeEngine.OpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WP = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Word;

/// <summary>
/// A live view over an inline image in a <see cref="WordDocument"/>.
/// </summary>
/// <remarks>
/// <para>
/// An embedded image is a package part (<c>word/media/image1.png</c>) plus a relationship from the
/// part that contains the drawing (<c>word/document.xml</c>, a header or a footer). This wrapper
/// exposes the image's size, content type and bytes; the relationship and the DrawingML markup stay
/// private.
/// </para>
/// <para>
/// This version models <strong>inline</strong> images whose picture is embedded through
/// <c>r:embed</c>. Floating images (<c>wp:anchor</c>), linked images (<c>r:link</c>) and other
/// graphics such as charts are preserved when a document is saved but are not returned by
/// <see cref="Paragraph.Images"/> and cannot be edited through this API.
/// </para>
/// <para>
/// This is a lightweight wrapper: reading it mutates nothing, and it must not be used after the
/// owning <see cref="WordDocument"/> has been disposed.
/// </para>
/// </remarks>
public sealed class Image
{
    private const double EmusPerPoint = 12700.0;
    private const string PictureNamespace = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    private readonly WP.Inline _inline;
    private readonly ImagePart _part;

    internal Image(W.Drawing drawing, WP.Inline inline, ImagePart part, string relationshipId)
    {
        Drawing = drawing;
        _inline = inline;
        _part = part;
        RelationshipId = relationshipId;
    }

    /// <summary>Gets the relationship id that links the drawing to its image part.</summary>
    public string RelationshipId { get; }

    /// <summary>Gets the image width in points, as stored in <c>wp:extent/@cx</c>.</summary>
    /// <exception cref="InvalidOfficeFileException">The stored extent is missing or not positive.</exception>
    public double WidthPoints => EmusToPoints(ExtentValue(horizontal: true));

    /// <summary>Gets the image height in points, as stored in <c>wp:extent/@cy</c>.</summary>
    /// <exception cref="InvalidOfficeFileException">The stored extent is missing or not positive.</exception>
    public double HeightPoints => EmusToPoints(ExtentValue(horizontal: false));

    /// <summary>Gets the MIME type of the image part, for example <c>image/png</c>.</summary>
    public string ContentType => _part.ContentType;

    /// <summary>Returns a copy of the bytes stored in the image part.</summary>
    /// <remarks>
    /// The returned array is a copy: changing it does not change the document. The image part is
    /// read-only for this API; use <see cref="Paragraph.AddImage(Stream, ImageFormat, double, double)"/>
    /// to add another image.
    /// </remarks>
    /// <exception cref="ObjectDisposedException">The owning document has been disposed.</exception>
    /// <exception cref="InvalidOfficeFileException">The image part cannot be read.</exception>
    public byte[] ReadAllBytes()
    {
        try
        {
            using Stream stream = _part.GetStream();
            using MemoryStream copy = new();
            stream.CopyTo(copy);
            return copy.ToArray();
        }
        catch (IOException exception)
        {
            throw new InvalidOfficeFileException("The image part cannot be read.", exception);
        }
    }

    /// <summary>Gets the drawing element that carries this image, for internal composition only.</summary>
    internal W.Drawing Drawing { get; }

    /// <summary>
    /// Creates an inline image part and the DrawingML that references it.
    /// </summary>
    /// <param name="owner">The part that will own the relationship (document, header or footer).</param>
    /// <param name="data">Image bytes to copy into the package.</param>
    /// <param name="format">The image format, which fixes the part's content type.</param>
    /// <param name="widthEmus">Image width in English Metric Units.</param>
    /// <param name="heightEmus">Image height in English Metric Units.</param>
    /// <param name="name">File name stored in the picture's non-visual properties.</param>
    internal static Image CreateInline(
        OpenXmlPart owner,
        Stream data,
        ImageFormat format,
        long widthEmus,
        long heightEmus,
        string name)
    {
        uint drawingId = NextDrawingId(owner);
        ImagePart part = AddImagePart(owner, format);
        try
        {
            part.FeedData(data);
        }
        catch
        {
            // Feeding untrusted or unreadable data must not leave an empty relationship behind.
            owner.DeletePart(part);
            throw;
        }

        string relationshipId = owner.GetIdOfPart(part);
        W.Drawing drawing = CreateDrawing(drawingId, name, relationshipId, widthEmus, heightEmus);
        return new Image(drawing, drawing.Inline!, part, relationshipId);
    }

    /// <summary>Collects the inline images that are direct children of a run.</summary>
    internal static void AddRunImages(W.Run run, List<Image> images)
    {
        foreach (W.Drawing drawing in run.Elements<W.Drawing>())
        {
            WP.Inline? inline = drawing.Inline;
            if (inline is null)
            {
                // A floating image (wp:anchor) is valid and preserved, but this version does not model it.
                continue;
            }

            PIC.Picture? picture = inline.Descendants<PIC.Picture>().FirstOrDefault();
            if (picture is null)
            {
                // Charts and other DrawingML graphics are valid and preserved, but not images here.
                continue;
            }

            A.Blip? blip = picture.BlipFill?.Blip;
            if (blip is null)
            {
                throw new InvalidOfficeFileException("An inline picture has no a:blip element.");
            }

            if (blip.Embed is null)
            {
                // A linked image (r:link) is an external relationship, which this version does not model.
                continue;
            }

            string relationshipId = blip.Embed.Value
                ?? throw new InvalidOfficeFileException("An inline picture has an empty image relationship id.");

            WP.Extent extent = inline.Extent
                ?? throw new InvalidOfficeFileException("An inline image has no wp:extent element.");
            long width = extent.Cx?.Value
                ?? throw new InvalidOfficeFileException("An inline image has no wp:extent/@cx attribute.");
            long height = extent.Cy?.Value
                ?? throw new InvalidOfficeFileException("An inline image has no wp:extent/@cy attribute.");
            if (width <= 0 || height <= 0)
            {
                throw new InvalidOfficeFileException("An inline image has a non-positive extent.");
            }

            OpenXmlPart owner = ResolvePart(drawing);
            if (!owner.TryGetPartById(relationshipId, out OpenXmlPart? related) || related is not ImagePart part)
            {
                throw new InvalidOfficeFileException(
                    $"An inline image references a relationship ('{relationshipId}') that is not an image part.");
            }

            images.Add(new Image(drawing, inline, part, relationshipId));
        }
    }

    /// <summary>Finds the part that owns the root element containing <paramref name="element"/>.</summary>
    internal static OpenXmlPart ResolvePart(OpenXmlElement element)
    {
        OpenXmlPartRootElement? root = element.Ancestors<OpenXmlPartRootElement>().FirstOrDefault();
        return root?.OpenXmlPart
            ?? throw new InvalidOfficeFileException("The image is not attached to a document part.");
    }

    /// <summary>Converts a positive number of points to English Metric Units.</summary>
    internal static long ToEmus(double points, string parameterName)
    {
        if (double.IsNaN(points) || double.IsInfinity(points) || points <= 0)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The image size must be a positive, finite number of points.");
        }

        const double LongLimit = 9223372036854775808.0; // 2^63, just above long.MaxValue.
        double emus = points * EmusPerPoint;
        if (!double.IsFinite(emus) || emus < 1 || emus >= LongLimit)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The image size is outside the range that Office Open XML can store.");
        }

        double rounded = Math.Round(emus, MidpointRounding.AwayFromZero);
        if (rounded >= LongLimit)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                points,
                "The image size is outside the range that Office Open XML can store.");
        }

        return (long)rounded;
    }

    /// <summary>Infers the image format from a file extension.</summary>
    internal static ImageFormat FromPath(string path)
    {
        string extension = Path.GetExtension(path).ToLowerInvariant();
        return extension switch
        {
            ".png" => ImageFormat.Png,
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            ".gif" => ImageFormat.Gif,
            ".bmp" => ImageFormat.Bmp,
            ".tif" or ".tiff" => ImageFormat.Tiff,
            ".emf" => ImageFormat.Emf,
            ".wmf" => ImageFormat.Wmf,
            _ => throw new ArgumentException(
                $"The file extension '{extension}' is not a supported image format. " +
                "Supported extensions: .png, .jpg, .jpeg, .gif, .bmp, .tif, .tiff, .emf and .wmf.",
                nameof(path)),
        };
    }

    /// <summary>Returns the conventional file name used for an image added from a stream.</summary>
    internal static string NameFor(ImageFormat format) => format switch
    {
        ImageFormat.Png => "image.png",
        ImageFormat.Jpeg => "image.jpg",
        ImageFormat.Gif => "image.gif",
        ImageFormat.Bmp => "image.bmp",
        ImageFormat.Tiff => "image.tiff",
        ImageFormat.Emf => "image.emf",
        ImageFormat.Wmf => "image.wmf",
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unknown image format."),
    };

    /// <summary>Validates that an image format value is defined.</summary>
    internal static void EnsureDefined(ImageFormat format, string parameterName)
    {
        if (!Enum.IsDefined(format))
        {
            throw new ArgumentOutOfRangeException(parameterName, format, "Unknown image format.");
        }
    }

    private long ExtentValue(bool horizontal)
    {
        WP.Extent? extent = _inline.Extent;
        long? value = horizontal ? extent?.Cx?.Value : extent?.Cy?.Value;
        if (value is null || value <= 0)
        {
            throw new InvalidOfficeFileException("An inline image has no positive wp:extent value.");
        }

        return value.Value;
    }

    private static double EmusToPoints(long emus) => emus / EmusPerPoint;

    private static W.Drawing CreateDrawing(
        uint id,
        string name,
        string relationshipId,
        long widthEmus,
        long heightEmus)
    {
        WP.Extent extent = new() { Cx = widthEmus, Cy = heightEmus };
        WP.EffectExtent effectExtent = new()
        {
            LeftEdge = 0L,
            TopEdge = 0L,
            RightEdge = 0L,
            BottomEdge = 0L,
        };
        WP.DocProperties docProperties = new() { Id = id, Name = $"Picture {id}" };
        WP.NonVisualGraphicFrameDrawingProperties frameProperties = new(
            new A.GraphicFrameLocks { NoChangeAspect = true });
        PIC.BlipFill blipFill = new(
            new A.Blip { Embed = relationshipId },
            new A.Stretch(new A.FillRectangle()));
        PIC.ShapeProperties shapeProperties = new(
            new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = widthEmus, Cy = heightEmus }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });
        PIC.Picture picture = new(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                new PIC.NonVisualPictureDrawingProperties()),
            blipFill,
            shapeProperties);
        A.Graphic graphic = new(
            new A.GraphicData(picture) { Uri = PictureNamespace });
        WP.Inline inline = new(extent, effectExtent, docProperties, frameProperties, graphic)
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        };

        return new W.Drawing(inline);
    }

    private static uint NextDrawingId(OpenXmlPart owner)
    {
        uint max = 0;
        foreach (OpenXmlPartRootElement root in WordRoots(owner))
        {
            foreach (WP.DocProperties properties in root.Descendants<WP.DocProperties>())
            {
                uint id = properties.Id?.Value ?? 0;
                if (id > max)
                {
                    max = id;
                }
            }
        }

        if (max == uint.MaxValue)
        {
            throw new InvalidOperationException("The document has no free drawing id.");
        }

        return max + 1;
    }

    private static IEnumerable<OpenXmlPartRootElement> WordRoots(OpenXmlPart owner)
    {
        if (owner.OpenXmlPackage is WordprocessingDocument word)
        {
            MainDocumentPart? main = word.MainDocumentPart;
            if (main is not null)
            {
                yield return OoxmlPartReader.Read("the main document part", () => main.Document);

                foreach (HeaderPart header in main.HeaderParts)
                {
                    yield return OoxmlPartReader.Read("the header part", () => header.Header);
                }

                foreach (FooterPart footer in main.FooterParts)
                {
                    yield return OoxmlPartReader.Read("the footer part", () => footer.Footer);
                }
            }

            yield break;
        }

        if (owner.RootElement is { } root)
        {
            yield return root;
        }
    }

    /// <summary>
    /// Adds an image part to a part that Office Open XML allows image relationships on.
    /// </summary>
    /// <remarks>
    /// The SDK exposes <c>AddImagePart</c> only for part types that declare the relationship, so the
    /// owning part is matched explicitly instead of being treated as a generic container.
    /// </remarks>
    private static ImagePart AddImagePart(OpenXmlPart owner, ImageFormat format)
    {
        PartTypeInfo type = ToOoxmlFormat(format);
        return owner switch
        {
            MainDocumentPart main => main.AddImagePart(type),
            HeaderPart header => header.AddImagePart(type),
            FooterPart footer => footer.AddImagePart(type),
            _ => throw new InvalidOperationException(
                $"Images cannot be added to a '{owner.GetType().Name}' part."),
        };
    }

    private static PartTypeInfo ToOoxmlFormat(ImageFormat format) => format switch
    {
        ImageFormat.Png => ImagePartType.Png,
        ImageFormat.Jpeg => ImagePartType.Jpeg,
        ImageFormat.Gif => ImagePartType.Gif,
        ImageFormat.Bmp => ImagePartType.Bmp,
        ImageFormat.Tiff => ImagePartType.Tiff,
        ImageFormat.Emf => ImagePartType.Emf,
        ImageFormat.Wmf => ImagePartType.Wmf,
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unknown image format."),
    };
}
