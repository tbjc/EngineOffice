using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Word;

public sealed class WordDocumentTests
{
    [Fact]
    public void Create_StartsEmpty()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Empty(document.Paragraphs);
        Assert.Equal(string.Empty, document.Text);
        Assert.Empty(document.Diagnostics);
    }

    [Fact]
    public void Create_SaveToStream_Reopen_RoundTripsParagraphText()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("Primera línea");
            document.AddParagraph("Segunda línea");
            bytes = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));

        Assert.Equal(2, reopened.Paragraphs.Count);
        Assert.Equal("Primera línea\nSegunda línea", reopened.Text);
        Assert.Empty(reopened.Diagnostics);
    }

    [Fact]
    public void Create_PreservesWhitespaceInParagraphText()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("  leading and trailing  ");
            bytes = SaveBytes.ToBytes(document);
        }

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));
        Assert.Equal("  leading and trailing  ", reopened.Text);
    }

    [Fact]
    public void AddParagraph_ReturnsEditableParagraph()
    {
        using WordDocument document = WordDocument.Create();
        Paragraph paragraph = document.AddParagraph("draft");

        Assert.Equal("draft", paragraph.Text);

        paragraph.AppendText(" final");
        Assert.Equal("draft final", paragraph.Text);

        paragraph.Text = "replaced";
        Assert.Equal("replaced", paragraph.Text);
        Assert.Equal("replaced", document.Text);
    }

    [Fact]
    public void AddParagraph_RejectsNull()
    {
        using WordDocument document = WordDocument.Create();
        Assert.Throws<ArgumentNullException>(() => document.AddParagraph(null!));
    }

    [Fact]
    public void Save_CanBeCalledRepeatedly_AndKeepsEveryEdit()
    {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("one");
        _ = SaveBytes.ToBytes(document);

        document.AddParagraph("two");
        byte[] bytes = SaveBytes.ToBytes(document);

        using WordDocument reopened = WordDocument.Open(TestPackage.ToStream(bytes));
        Assert.Equal("one\ntwo", reopened.Text);
    }

    [Fact]
    public void Save_ToFile_ThenOpen_RoundTrips()
    {
        string directory = TestData.NewTemporaryDirectory();
        string path = Path.Combine(directory, "document.docx");
        try
        {
            using (WordDocument document = WordDocument.Create())
            {
                document.AddParagraph("on disk");
                document.Save(path);
            }

            Assert.True(File.Exists(path));

            using (WordDocument reopened = WordDocument.Open(path))
            {
                Assert.Equal("on disk", reopened.Text);
            }

            // Saving over the file that was opened must succeed.
            using (WordDocument again = WordDocument.Open(path))
            {
                again.AddParagraph("second pass");
                again.Save(path);
            }

            using (WordDocument third = WordDocument.Open(path))
            {
                Assert.Equal("on disk\nsecond pass", third.Text);
            }
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Open_FromStream_DisposesSourceStreamByDefault()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("x");
            bytes = SaveBytes.ToBytes(document);
        }

        MemoryStream source = TestPackage.ToStream(bytes);
        using (WordDocument.Open(source))
        {
        }

        Assert.False(source.CanRead);
    }

    [Fact]
    public void Open_FromStream_WithLeaveOpen_KeepsSourceStreamUsable()
    {
        byte[] bytes;
        using (WordDocument document = WordDocument.Create())
        {
            document.AddParagraph("x");
            bytes = SaveBytes.ToBytes(document);
        }

        using MemoryStream source = TestPackage.ToStream(bytes);
        using WordDocument document2 = WordDocument.Open(source, leaveOpen: true);

        Assert.True(source.CanRead);
        Assert.Equal("x", document2.Text);
    }

    [Fact]
    public void Open_NonexistentFile_ThrowsFileNotFound()
    {
        string path = Path.Combine(TestData.NewTemporaryDirectory(), "missing.docx");
        Assert.Throws<FileNotFoundException>(() => WordDocument.Open(path));
    }

    [Fact]
    public void DisposedDocument_ThrowsObjectDisposedException()
    {
        WordDocument document = WordDocument.Create();
        document.AddParagraph("x");
        document.Dispose();

        using MemoryStream target = new();
        Assert.Throws<ObjectDisposedException>(() => _ = document.Paragraphs);
        Assert.Throws<ObjectDisposedException>(() => document.AddParagraph("y"));
        Assert.Throws<ObjectDisposedException>(() => document.Save(target));
        Assert.Throws<ObjectDisposedException>(() => _ = document.Text);
    }

    [Fact]
    public void Save_ToNonWritableStream_ThrowsArgumentException()
    {
        using WordDocument document = WordDocument.Create();
        using MemoryStream readOnly = new([], writable: false);
        Assert.Throws<ArgumentException>(() => document.Save(readOnly));
    }

    [Fact]
    public void Save_RejectsInvalidArguments()
    {
        using WordDocument document = WordDocument.Create();
        Assert.Throws<ArgumentNullException>(() => document.Save((Stream)null!));
        Assert.Throws<ArgumentException>(() => document.Save(string.Empty));
    }

    [Fact]
    public void Open_RejectsNullStream() => Assert.Throws<ArgumentNullException>(() => WordDocument.Open((Stream)null!));
}
