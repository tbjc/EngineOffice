using OfficeEngine;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;

namespace OfficeEngine.Tests.Excel;

/// <summary>Notes of a worksheet: the comments part, its authors and the legacy drawing that goes with it.</summary>
public sealed class ExcelCommentTests
{
    [Fact]
    public void AddComment_WritesTheCommentPartAndTheDrawing()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            ExcelComment comment = sheet.AddComment("A1", "Revisar el total", "Ana");

            Assert.Equal("A1", comment.Reference);
            Assert.Equal("Revisar el total", comment.Text);
            Assert.Equal("Ana", comment.Author);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.True(TestPackage.HasEntry(bytes, "xl/comments1.xml"));
        Assert.Contains(
            TestPackage.EntryNames(bytes),
            name => name.EndsWith(".vml", StringComparison.Ordinal));

        IReadOnlyList<(string Reference, string Text, string Author)> comments =
            OoxmlValidation.ReadXlsxComments(bytes, "Sheet1");
        Assert.Single(comments);
        Assert.Equal("A1", comments[0].Reference);
        Assert.Equal("Revisar el total", comments[0].Text);
        Assert.Equal("Ana", comments[0].Author);

        using ExcelWorkbook reopened = ExcelWorkbook.Open(TestPackage.ToStream(bytes));
        Assert.Equal("Revisar el total", reopened.Worksheets[0].Comments[0].Text);
    }

    [Fact]
    public void AddComment_ForTheSameCell_ReplacesTheNote()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.AddComment("A1", "Primero");

        sheet.AddComment("A1", "Segundo");

        Assert.Single(sheet.Comments);
        Assert.Equal("Segundo", sheet.Comments[0].Text);
    }

    [Fact]
    public void AddComment_WithSeveralAuthors_KeepsTheAuthorList()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AddComment("A1", "Una", "Ana");
            sheet.AddComment("A2", "Dos", "Luis");

            Assert.Equal(2, sheet.Comments.Count);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        IReadOnlyList<(string Reference, string Text, string Author)> comments =
            OoxmlValidation.ReadXlsxComments(bytes, "Sheet1");
        Assert.Equal("Ana", comments[0].Author);
        Assert.Equal("Luis", comments[1].Author);
    }

    [Fact]
    public void AddComment_WithInvalidValues_Throws()
    {
        using ExcelWorkbook workbook = ExcelWorkbook.Create();
        Worksheet sheet = workbook.Worksheets[0];

        Assert.Throws<ArgumentException>(() => sheet.AddComment("B2:A1", "Texto"));
        Assert.Throws<ArgumentException>(() => sheet.AddComment("A1", string.Empty));
        Assert.Throws<ArgumentException>(() => sheet.AddComment("A1", "Texto", string.Empty));
        Assert.Empty(sheet.Comments);
    }

    [Fact]
    public void RemoveComment_RemovesTheNote()
    {
        byte[] bytes;
        using (ExcelWorkbook workbook = ExcelWorkbook.Create())
        {
            Worksheet sheet = workbook.Worksheets[0];
            sheet.AddComment("A1", "Revisar");

            Assert.True(sheet.RemoveComment("A1"));
            Assert.False(sheet.RemoveComment("A1"));
            Assert.Empty(sheet.Comments);

            bytes = SaveBytes.ToBytes(workbook);
        }

        Assert.Empty(OoxmlValidation.ValidateXlsx(bytes));
        Assert.Empty(OoxmlValidation.ReadXlsxComments(bytes, "Sheet1"));
    }

    [Fact]
    public void Comments_ReadsTheNotesOfAnotherWorkbook()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithComments(
            ("A1", "Primera nota", "Ana"),
            ("B2", "Segunda nota", "Luis"));

        using ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture));

        Assert.Equal(2, workbook.Worksheets[0].Comments.Count);
        Assert.Equal("Primera nota", workbook.Worksheets[0].Comments[0].Text);
        Assert.Equal("Luis", workbook.Worksheets[0].Comments[1].Author);
    }

    [Fact]
    public void Comments_ArePreservedWhenTheWorkbookIsSaved()
    {
        byte[] fixture = OoxmlValidation.CreateXlsxWithComments(("A1", "Nota", "Ana"));

        byte[] produced;
        using (ExcelWorkbook workbook = ExcelWorkbook.Open(TestPackage.ToStream(fixture)))
        {
            produced = SaveBytes.ToBytes(workbook);
        }

        Assert.Single(OoxmlValidation.ReadXlsxComments(produced, "Hoja1"));
        Assert.Empty(OoxmlValidation.ValidateXlsx(produced));
    }

    [Fact]
    public void DisposedWorkbook_CommentsThrowObjectDisposedException()
    {
        ExcelWorkbook workbook = ExcelWorkbook.Create();
        workbook.Dispose();

        Assert.Throws<ObjectDisposedException>(() => _ = workbook.Worksheets[0].Comments);
    }
}
