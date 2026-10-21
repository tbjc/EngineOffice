using System;
using System.IO;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.10: core document properties. A created document starts without them, values set
/// through the API must survive a round-trip, and extended/custom properties that this version does
/// not model must not be dropped.
/// </summary>
public sealed class WordPropertiesTests
{
    private static readonly DateTime CreatedInstant = new(2026, 1, 2, 3, 4, 5, DateTimeKind.Utc);
    private static readonly DateTime ModifiedInstant = new(2026, 2, 3, 4, 5, 6, DateTimeKind.Utc);
    private static readonly DateTime PrintedInstant = new(2026, 3, 4, 5, 6, 7, DateTimeKind.Utc);

    [Fact]
    public void Create_HasNoCorePropertiesAndNoCorePart()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Null(document.Properties.Title);
        Assert.Null(document.Properties.Subject);
        Assert.Null(document.Properties.Creator);
        Assert.Null(document.Properties.Keywords);
        Assert.Null(document.Properties.Description);
        Assert.Null(document.Properties.LastModifiedBy);
        Assert.Null(document.Properties.Revision);
        Assert.Null(document.Properties.LastPrinted);
        Assert.Null(document.Properties.Created);
        Assert.Null(document.Properties.Modified);
        Assert.Null(document.Properties.Category);
        Assert.Null(document.Properties.ContentStatus);
        Assert.Null(document.Properties.Identifier);
        Assert.Null(document.Properties.Language);
        Assert.Null(document.Properties.Version);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.False(TestPackage.HasEntry(produced, "docProps/core.xml"));
    }

    [Fact]
    public void Properties_ClearingMissingValues_DoesNotCreateTheCorePart()
    {
        using WordDocument document = WordDocument.Create();

        document.Properties.Title = null;
        document.Properties.Created = null;

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.False(TestPackage.HasEntry(produced, "docProps/core.xml"));
    }

    [Fact]
    public void Properties_AllCoreValues_RoundTripThroughAFile()
    {
        string directory = TestData.NewTemporaryDirectory();
        string documentPath = Path.Combine(directory, "properties.docx");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                SetAll(document.Properties);
                document.Save(documentPath);
            }

            using (WordDocument reopened = WordDocument.Open(documentPath))
            {
                AssertAll(reopened.Properties);
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Properties_ExistingValues_AreEditableAndRoundTripThroughAStream()
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

            document.Properties.Title = "Nuevo";
            document.Properties.Subject = "Tema";
            document.Properties.Keywords = "uno; dos";
            document.Properties.Description = "Descripcion";
            document.Properties.LastModifiedBy = "Editor";
            document.Properties.Revision = "7";
            document.Properties.Modified = ModifiedInstant;
            document.Properties.LastPrinted = PrintedInstant;
            document.Properties.Category = "Categoria";
            document.Properties.ContentStatus = "Borrador";
            document.Properties.Identifier = "id-1";
            document.Properties.Language = "es-ES";
            document.Properties.Version = "1.2";

            produced = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal("Nuevo", reopened.Properties.Title);
        Assert.Equal("Tema", reopened.Properties.Subject);
        Assert.Equal("Autor", reopened.Properties.Creator);
        Assert.Equal("uno; dos", reopened.Properties.Keywords);
        Assert.Equal("Descripcion", reopened.Properties.Description);
        Assert.Equal("Editor", reopened.Properties.LastModifiedBy);
        Assert.Equal("7", reopened.Properties.Revision);
        AssertInstant(ModifiedInstant, reopened.Properties.Modified);
        AssertInstant(PrintedInstant, reopened.Properties.LastPrinted);
        Assert.Equal("Categoria", reopened.Properties.Category);
        Assert.Equal("Borrador", reopened.Properties.ContentStatus);
        Assert.Equal("id-1", reopened.Properties.Identifier);
        Assert.Equal("es-ES", reopened.Properties.Language);
        Assert.Equal("1.2", reopened.Properties.Version);
    }

    [Fact]
    public void Properties_CanBeCleared()
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
            document.Properties.Title = null;
            document.Properties.Creator = null;
            document.Properties.Created = null;
            produced = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Null(reopened.Properties.Title);
        Assert.Null(reopened.Properties.Creator);
        Assert.Null(reopened.Properties.Created);

        IReadOnlyDictionary<string, string?> stored = OoxmlValidation.ReadDocxCoreProperties(produced);
        Assert.DoesNotContain("title", stored);
        Assert.DoesNotContain("creator", stored);
        Assert.DoesNotContain("created", stored);
    }

    [Fact]
    public void Properties_PreserveExtendedAndCustomProperties()
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
            document.Properties.Title = "Nuevo";
            produced = SaveBytes.ToBytes(document);
        }

        Assert.Equal("Nuevo", OoxmlValidation.ReadDocxCoreProperties(produced)["title"]);
        Assert.Equal("Acme", OoxmlValidation.ReadDocxCompany(produced));
        Assert.Equal("valor", OoxmlValidation.ReadDocxCustomPropertyValue(produced, "CustomName"));
    }

    [Fact]
    public void Properties_AllowTabsAndLineBreaks()
    {
        const string description = "linea uno\nlinea dos\tfinal";
        byte[] produced;

        using (WordDocument document = WordDocument.Create())
        {
            document.Properties.Description = description;
            produced = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(produced));
        Assert.Equal(description, reopened.Properties.Description);
    }

    [Fact]
    public void Properties_RejectCharactersXmlCannotStore()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Throws<ArgumentException>(() => document.Properties.Title = "a\u0000b");
        Assert.Throws<ArgumentException>(() => document.Properties.Description = "\uD800");
        Assert.Null(document.Properties.Title);
        Assert.Null(document.Properties.Description);

        byte[] produced = SaveBytes.ToBytes(document);
        Assert.False(TestPackage.HasEntry(produced, "docProps/core.xml"));
    }

    [Fact]
    public void Properties_ReportAMalformedCorePartWithATypedError()
    {
        byte[] fixture = OoxmlValidation.CreateDocxWithMalformedCorePropertyDate("Titulo");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(fixture));

        Assert.Throws<InvalidOfficeFileException>(() => document.Properties.Created);
    }

    [Fact]
    public void Properties_AfterDispose_ThrowObjectDisposedException()
    {
        WordDocument document = WordDocument.Create();
        document.Dispose();

        Assert.Throws<ObjectDisposedException>(() => document.Properties.Title);
        Assert.Throws<ObjectDisposedException>(() => document.Properties.Title = "x");
        Assert.Throws<ObjectDisposedException>(() => document.Properties.Created);
        Assert.Throws<ObjectDisposedException>(() => document.Properties.Created = CreatedInstant);
    }

    private static void SetAll(DocumentProperties properties)
    {
        properties.Title = "Titulo";
        properties.Subject = "Tema";
        properties.Creator = "Autor";
        properties.Keywords = "uno; dos";
        properties.Description = "Descripcion";
        properties.LastModifiedBy = "Editor";
        properties.Revision = "7";
        properties.LastPrinted = PrintedInstant;
        properties.Created = CreatedInstant;
        properties.Modified = ModifiedInstant;
        properties.Category = "Categoria";
        properties.ContentStatus = "Borrador";
        properties.Identifier = "id-1";
        properties.Language = "es-ES";
        properties.Version = "1.2";
    }

    private static void AssertAll(DocumentProperties properties)
    {
        Assert.Equal("Titulo", properties.Title);
        Assert.Equal("Tema", properties.Subject);
        Assert.Equal("Autor", properties.Creator);
        Assert.Equal("uno; dos", properties.Keywords);
        Assert.Equal("Descripcion", properties.Description);
        Assert.Equal("Editor", properties.LastModifiedBy);
        Assert.Equal("7", properties.Revision);
        AssertInstant(PrintedInstant, properties.LastPrinted);
        AssertInstant(CreatedInstant, properties.Created);
        AssertInstant(ModifiedInstant, properties.Modified);
        Assert.Equal("Categoria", properties.Category);
        Assert.Equal("Borrador", properties.ContentStatus);
        Assert.Equal("id-1", properties.Identifier);
        Assert.Equal("es-ES", properties.Language);
        Assert.Equal("1.2", properties.Version);
    }

    private static void AssertInstant(DateTime expected, DateTime? actual)
    {
        Assert.True(actual.HasValue);
        Assert.Equal(expected.ToUniversalTime(), actual.Value.ToUniversalTime());
    }
}
