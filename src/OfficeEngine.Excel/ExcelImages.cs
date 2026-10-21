using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeEngine.Excel;

/// <summary>Builds and reads the drawing that holds the images of a worksheet.</summary>
internal static class ExcelImages
{
    /// <summary>The number of English Metric Units in a point.</summary>
    internal const long EmusPerPoint = 12700L;

    internal static Xdr.OneCellAnchor BuildAnchor(
        uint id,
        string name,
        string relationshipId,
        int column,
        int row,
        long width,
        long height)
    {
        Xdr.FromMarker from = new(
            new Xdr.ColumnId(column.ToString(CultureInfo.InvariantCulture)),
            new Xdr.ColumnOffset("0"),
            new Xdr.RowId(row.ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowOffset("0"));

        Xdr.Extent extent = new() { Cx = width, Cy = height };

        A.Blip blip = new() { Embed = relationshipId };
        Xdr.BlipFill blipFill = new(blip, new A.Stretch(new A.FillRectangle()));
        A.Transform2D transform = new(
            new A.Offset { X = 0L, Y = 0L },
            new A.Extents { Cx = width, Cy = height });
        A.PresetGeometry geometry = new(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle };
        Xdr.ShapeProperties shapeProperties = new(transform, geometry);

        Xdr.NonVisualDrawingProperties properties = new() { Id = id, Name = name };
        Xdr.NonVisualPictureProperties nonVisual = new(
            properties,
            new Xdr.NonVisualPictureDrawingProperties());
        Xdr.Picture picture = new(nonVisual, blipFill, shapeProperties);

        return new Xdr.OneCellAnchor(from, extent, picture, new Xdr.ClientData());
    }

    internal static List<ExcelImage> Read(DrawingsPart drawingPart)
    {
        List<ExcelImage> images = [];
        Xdr.WorksheetDrawing? drawing = drawingPart.WorksheetDrawing;
        if (drawing is null)
        {
            return images;
        }

        foreach (Xdr.OneCellAnchor anchor in drawing.Elements<Xdr.OneCellAnchor>())
        {
            Xdr.Picture? picture = anchor.GetFirstChild<Xdr.Picture>();
            Xdr.FromMarker? from = anchor.GetFirstChild<Xdr.FromMarker>();
            Xdr.Extent? extent = anchor.GetFirstChild<Xdr.Extent>();
            if (picture is null || from is null)
            {
                continue;
            }

            string name = picture.NonVisualPictureProperties?
                .NonVisualDrawingProperties?.Name?.Value ?? string.Empty;
            int column = ParseIndex(from.ColumnId?.Text) + 1;
            int row = ParseIndex(from.RowId?.Text) + 1;
            double width = extent?.Cx?.Value is long cx ? cx / (double)EmusPerPoint : 0d;
            double height = extent?.Cy?.Value is long cy ? cy / (double)EmusPerPoint : 0d;
            string mediaType = MediaTypeOf(drawingPart, picture);
            images.Add(new ExcelImage(name, $"{CellReference.ToReference(column, row)}", width, height, mediaType));
        }

        return images;
    }

    internal static string MediaTypeOf(DrawingsPart drawingPart, Xdr.Picture picture)
    {
        string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
        if (relationshipId is null)
        {
            return string.Empty;
        }

        ImagePart? imagePart = drawingPart.GetPartById(relationshipId) as ImagePart;
        return imagePart?.ContentType ?? string.Empty;
    }

    private static int ParseIndex(string? text)
        => int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : 0;
}
