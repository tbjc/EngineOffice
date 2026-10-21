using System;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.10 interoperability checks: generated core properties must be schema-valid, use the
/// standard OOXML namespaces and be readable by the Open XML SDK. Fixtures from the SDK must be
/// readable and editable without losing parts that OfficeEngine does not model.
/// </summary>
public sealed class WordPropertiesInteropTests
{
    private static readonly DateTime CreatedInstant = new(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc);
    private static readonly DateTime ModifiedInstant = new(2026, 2, 3, 4, 5, 6, DateTimeKind.Utc);

    [Fact]
    public void ProducedDocument_WithCoreProperties_IsValidAndReadableByTheSdk()
    {
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            document.Properties.Title = "Titulo";
            document.Properties.Subject = "Tema";
            document.Properties.Creator = "Autor";
            document.Properties.Keywords = "uno; dos";
            document.Properties.Description = "Descripcion";
            document.Properties.LastModifiedBy = "Editor";
            document.Properties.Revision = "7";
            document.Properties.LastPrinted = ModifiedInstant;
            document.Properties.Created = CreatedInstant;
            document.Properties.Modified = ModifiedInstant;
            document.Properties.Category = "Categoria";
            document.Properties.ContentStatus = "Borrador";
            document.Properties.Identifier = "id-1";
            document.Properties.Language = "es-ES";
            document.Properties.Version = "1.2";
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));

        using MemoryStream stream = new(produced, writable: false);
        using WordprocessingDocument sdk = WordprocessingDocument.Open(stream, isEditable: false);
        CoreFilePropertiesPart part = Assert.IsType<CoreFilePropertiesPart>(sdk.CoreFilePropertiesPart);
        Assert.Equal("/docProps/core.xml", part.Uri.ToString());
        Assert.Equal("Titulo", part.CoreFileProperties.Title);
        Assert.Equal("Tema", part.CoreFileProperties.Subject);
        Assert.Equal("Autor", part.CoreFileProperties.Creator);
        Assert.Equal("uno; dos", part.CoreFileProperties.Keywords);
        Assert.Equal("Descripcion", part.CoreFileProperties.Description);
        Assert.Equal("Editor", part.CoreFileProperties.LastModifiedBy);
        Assert.Equal("7", part.CoreFileProperties.Revision);
        AssertInstant(ModifiedInstant, part.CoreFileProperties.LastPrinted);
        AssertInstant(CreatedInstant, part.CoreFileProperties.Created);
        AssertInstant(ModifiedInstant, part.CoreFileProperties.Modified);
        Assert.Equal("Categoria", part.CoreFileProperties.Category);
        Assert.Equal("Borrador", part.CoreFileProperties.ContentStatus);
        Assert.Equal("id-1", part.CoreFileProperties.Identifier);
        Assert.Equal("es-ES", part.CoreFileProperties.Language);
        Assert.Equal("1.2", part.CoreFileProperties.Version);
    }

    [Fact]
    public void ProducedDocument_CoreProperties_UseTheExpectedElementsAndNamespaces()
    {
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            document.Properties.Title = "Titulo";
            document.Properties.Subject = "Tema";
            document.Properties.Creator = "Autor";
            document.Properties.Keywords = "uno; dos";
            document.Properties.Description = "Descripcion";
            document.Properties.LastModifiedBy = "Editor";
            document.Properties.Revision = "7";
            document.Properties.LastPrinted = ModifiedInstant;
            document.Properties.Created = CreatedInstant;
            document.Properties.Modified = ModifiedInstant;
            document.Properties.Category = "Categoria";
            document.Properties.ContentStatus = "Borrador";
            document.Properties.Identifier = "id-1";
            document.Properties.Language = "es-ES";
            document.Properties.Version = "1.2";
            produced = SaveBytes.ToBytes(document);
        }

        XDocument core = XDocument.Parse(OoxmlValidation.ReadDocxCorePropertiesXml(produced));
        XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        XNamespace dc = "http://purl.org/dc/elements/1.1/";
        XNamespace dcterms = "http://purl.org/dc/terms/";
        XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";
        XElement root = core.Root!;

        Assert.Equal(cp + "coreProperties", root.Name);
        Assert.Equal("Titulo", (string?)root.Element(dc + "title"));
        Assert.Equal("Tema", (string?)root.Element(dc + "subject"));
        Assert.Equal("Autor", (string?)root.Element(dc + "creator"));
        Assert.Equal("uno; dos", (string?)root.Element(cp + "keywords"));
        Assert.Equal("Descripcion", (string?)root.Element(dc + "description"));
        Assert.Equal("Editor", (string?)root.Element(cp + "lastModifiedBy"));
        Assert.Equal("7", (string?)root.Element(cp + "revision"));
        Assert.NotNull(root.Element(cp + "lastPrinted"));
        Assert.NotNull(root.Element(dcterms + "created"));
        Assert.NotNull(root.Element(dcterms + "modified"));
        Assert.Equal("Categoria", (string?)root.Element(cp + "category"));
        Assert.Equal("Borrador", (string?)root.Element(cp + "contentStatus"));
        Assert.Equal("id-1", (string?)root.Element(dc + "identifier"));
        Assert.Equal("es-ES", (string?)root.Element(dc + "language"));
        Assert.Equal("1.2", (string?)root.Element(cp + "version"));

        Assert.Equal("dcterms:W3CDTF", (string?)root.Element(dcterms + "created")?.Attribute(xsi + "type"));
        Assert.Equal("dcterms:W3CDTF", (string?)root.Element(dcterms + "modified")?.Attribute(xsi + "type"));
    }

    [Fact]
    public void OpenedDocument_WithCoreProperties_EditsThemAndPreservesUnmodelledParts()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithDocumentProperties(
            "Viejo",
            "Autor",
            CreatedInstant,
            "Acme",
            "valor");
        byte[] produced;

        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            Assert.Equal("Viejo", document.Properties.Title);
            Assert.Equal("Autor", document.Properties.Creator);
            AssertInstant(CreatedInstant, document.Properties.Created);
            Assert.Equal("Acme", OoxmlValidation.ReadDocxCompany(fixture));
            Assert.Equal("valor", OoxmlValidation.ReadDocxCustomPropertyValue(fixture, "CustomName"));

            document.Properties.Title = "Nuevo";
            document.Properties.Modified = ModifiedInstant;
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.Equal("Nuevo", OoxmlValidation.ReadDocxCoreProperties(produced)["title"]);
        Assert.Equal("Acme", OoxmlValidation.ReadDocxCompany(produced));
        Assert.Equal("valor", OoxmlValidation.ReadDocxCustomPropertyValue(produced, "CustomName"));

        using MemoryStream stream = new(produced, writable: false);
        using WordprocessingDocument sdk = WordprocessingDocument.Open(stream, isEditable: false);
        Assert.Equal("Nuevo", sdk.CoreFilePropertiesPart?.CoreFileProperties.Title);
        AssertInstant(ModifiedInstant, sdk.CoreFilePropertiesPart?.CoreFileProperties.Modified);
    }

    [Fact]
    public void OpenedDocument_WithoutCoreProperties_CreatesTheConventionalCorePart()
    {
        byte[] fixture;
        using (WordDocument created = WordDocument.Create())
        {
            fixture = SaveBytes.ToBytes(created);
        }

        byte[] produced;
        using (WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture)))
        {
            document.Properties.Creator = "Autor";
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Empty(OoxmlValidation.ValidateDocx(produced));
        Assert.True(TestPackage.HasEntry(produced, "docProps/core.xml"));
        Assert.Equal("Autor", OoxmlValidation.ReadDocxCoreProperties(produced)["creator"]);

        using MemoryStream stream = new(produced, writable: false);
        using WordprocessingDocument sdk = WordprocessingDocument.Open(stream, isEditable: false);
        Assert.Equal("/docProps/core.xml", sdk.CoreFilePropertiesPart?.Uri.ToString());
        Assert.Equal("Autor", sdk.CoreFilePropertiesPart?.CoreFileProperties.Creator);
    }

    private static void AssertInstant(DateTime expected, DateTime? actual)
    {
        Assert.True(actual.HasValue);
        Assert.Equal(expected.ToUniversalTime(), actual.Value.ToUniversalTime());
    }
}
