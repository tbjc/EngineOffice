using OfficeEngine.Diagnostics;
using OfficeEngine.Excel;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;

namespace OfficeEngine.Tests.Documentation;

/// <summary>
/// Runs the documented usage flow so the examples in README.md cannot drift from the real API.
/// </summary>
public sealed class ReadmeExamplesTests
{
    [Fact]
    public void Word_CreateEditSaveAndInspectDiagnostics()
    {
        string directory = TestData.NewTemporaryDirectory();
        string input = Path.Combine(directory, "informe.docx");
        string output = Path.Combine(directory, "salida.docx");
        try
        {
            // Word: crear, escribir y guardar.
            using (WordDocument document = WordDocument.Create())
            {
                document.AddParagraph("Informe mensual");

                // Un párrafo puede tener varios runs, cada uno con su propio formato de carácter.
                Paragraph summary = document.AddParagraph(string.Empty);
                Run amount = summary.AddRun("Cifras al alza");
                amount.Bold = true;
                amount.Underline = UnderlineStyle.SingleLine;
                amount.FontSizePoints = 11.0;
                amount.Color = new RgbColor(0x1F, 0x49, 0x7D);
                summary.AddRun(" frente al mes anterior.");

                // El formato de párrafo se modela en el propio párrafo.
                summary.Alignment = ParagraphAlignment.Center;
                summary.LeftIndentPoints = 18.0;
                summary.SpacingAfterPoints = 6.0;

                // Las tabulaciones pertenecen al párrafo; se invocan con Run.AddTab.
                Paragraph prices = document.AddParagraph("Precio");
                prices.AddTabStop(new TabStop(72.0, TabStopAlignment.Decimal, TabStopLeader.Dot));
                prices.AddRun(string.Empty).AddTab().AddText("1.234,56");

                // Un párrafo puede nombrar un estilo de párrafo y un run uno de carácter.
                summary.StyleId = "Heading1";
                amount.StyleId = "Emphasis";

                // La sección define el tamaño/orientación de página y los márgenes.
                document.Sections[0].PageSize = PageSize.A4;
                document.Sections[0].Margins = new PageMargins(72.0, 72.0, 72.0, 72.0);

                // Un encabezado o pie es una parte de la sección.
                document.Sections[0].AddHeader(HeaderFooterType.Primary).Paragraphs[0].Text = "Informe mensual";

                // El pie lleva el número de página: el campo se escribe, no se calcula.
                HeaderFooter footer = document.Sections[0].AddFooter(HeaderFooterType.Primary);
                footer.Paragraphs[0].AddRun("Página ");
                footer.Paragraphs[0].AddField("PAGE", "1");

                // Se crea una lista (numerada o con viñetas) con `document.AddList` y se aplica a un
                // párrafo con `Paragraph.ApplyList`; el nivel es opcional (0..8).
                WordList items = document.AddList(WordListFormat.Bullet);
                document.AddParagraph("Primer punto").ApplyList(items);
                document.AddParagraph("Segundo punto").ApplyList(items);

                // Se crea una tabla con `document.AddTable(filas, columnas)`; cada celda empieza con
                // un párrafo vacío que se rellena con `TableCell.Text`.
                Table totals = document.AddTable(2, 2);
                totals.Rows[0].Cells[0].Text = "Concepto";
                totals.Rows[0].Cells[1].Text = "Importe";
                totals.Rows[1].Cells[0].Text = "Total";
                totals.Rows[1].Cells[1].Text = "1.234,56";

                document.Save(input);
            }

            // Word: abrir, editar y guardar en otro archivo.
            using (WordDocument loaded = WordDocument.Open(input))
            {
                loaded.Paragraphs[0].Text = "Informe final";

                // Leer el formato de carácter de un run.
                Run amount = loaded.Paragraphs[1].Runs[0];
                Assert.True(amount.Bold);
                Assert.Equal(UnderlineStyle.SingleLine, amount.Underline);
                Assert.Equal(11.0, amount.FontSizePoints!.Value);
                Assert.Equal(new RgbColor(0x1F, 0x49, 0x7D), amount.Color!.Value);
                Assert.Equal("Cifras al alza frente al mes anterior.", loaded.Paragraphs[1].Text);

                // Leer el formato de párrafo.
                Assert.Equal(ParagraphAlignment.Center, loaded.Paragraphs[1].Alignment);
                Assert.Equal(18.0, loaded.Paragraphs[1].LeftIndentPoints!.Value);
                Assert.Equal(6.0, loaded.Paragraphs[1].SpacingAfterPoints!.Value);

                // Leer las tabulaciones de un párrafo y el contenido real de un run.
                TabStop stop = loaded.Paragraphs[2].TabStops[0];
                Assert.Equal(72.0, stop.PositionPoints);
                Assert.Equal(TabStopAlignment.Decimal, stop.Alignment);
                Assert.Equal(TabStopLeader.Dot, stop.Leader);

                // Las referencias de estilo se conservan tal cual (el estilo puede estar definido o no).
                Assert.Equal("Heading1", loaded.Paragraphs[1].StyleId);
                Assert.Equal("Emphasis", loaded.Paragraphs[1].Runs[0].StyleId);

                // La sección y su tamaño/márgenes también se conservan.
                Assert.Equal(PageSize.A4, loaded.Sections[0].PageSize);
                Assert.Equal(new PageMargins(72.0, 72.0, 72.0, 72.0), loaded.Sections[0].Margins);

                // El encabezado y el pie de la sección se leen por tipo.
                Assert.Equal("Informe mensual", loaded.Sections[0].Header(HeaderFooterType.Primary)!.Text);
                Assert.Equal("Página 1", loaded.Sections[0].Footer(HeaderFooterType.Primary)!.Text);

                // Buscar y reemplazar en el cuerpo.
                Assert.Equal(1, loaded.ReplaceText("frente al mes anterior", "del mes pasado"));
                Assert.Single(loaded.FindText("del mes pasado"));

                // Leer la lista de un párrafo y el formato del nivel al que pertenece.
                Assert.Equal(1, loaded.Paragraphs[3].ListId);
                Assert.Equal(0, loaded.Paragraphs[3].ListLevel);
                Assert.Equal(WordListFormat.Bullet, loaded.FindList(1)!.Levels[0].Format);

                // Leer una tabla y añadir una fila con una celda por columna.
                Table totals = loaded.Tables[0];
                Assert.Equal(2, totals.RowCount);
                Assert.Equal(2, totals.ColumnCount);
                Assert.Equal("Concepto", totals.Rows[0].Cells[0].Text);
                TableRow note = totals.AddRow();
                note.Cells[0].Text = "Nota";
                note.Cells[1].Text = "0";

                loaded.Save(output);

                foreach (OfficeEngineDiagnostic diagnostic in loaded.Diagnostics)
                {
                    Assert.NotNull(diagnostic.Code);
                }
            }

            using WordDocument result = WordDocument.Open(output);
            Assert.Equal(
                "Informe final\nCifras al alza del mes pasado.\nPrecio1.234,56\nPrimer punto\n" +
                "Segundo punto\nConcepto\tImporte\nTotal\t1.234,56\nNota\t0",
                result.Text);

            // Run.Text solo concatena w:t; la tabulación está en Run.Content.
            Assert.Collection(
                result.Paragraphs[2].Runs[1].Content,
                item => Assert.IsType<RunTab>(item),
                item => Assert.Equal("1.234,56", Assert.IsType<RunText>(item).Value));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Excel_CreateWriteCellsAndSave()
    {
        string directory = TestData.NewTemporaryDirectory();
        string path = Path.Combine(directory, "resumen.xlsx");
        try
        {
            using (ExcelWorkbook workbook = ExcelWorkbook.Create())
            {
                Worksheet sheet = workbook.Worksheets.Add("Resumen");
                sheet["A1"].Value = "Concepto";
                sheet["B1"].Value = 125;
                sheet["B2"].Value = 145.5;
                sheet["C1"].Formula = "=B1*1.16";
                sheet.AddNamedRange("Tabla", "A1:B2");
                sheet.AddDataValidation(ExcelDataValidation.List("C1:C10", ["bajo", "medio", "alto"]));
                sheet.AutoFilter = new ExcelAutoFilter("D1:E10");
                sheet.AddTable("Resumen", "A1:C10");
                workbook.Save(path);
            }

            using ExcelWorkbook reopened = ExcelWorkbook.Open(path);
            Worksheet resumen = reopened.Worksheets["Resumen"];
            Assert.Equal("Concepto", resumen["A1"].Value);
            Assert.Equal(125.0, resumen["B1"].Value);
            Assert.Equal(145.5, resumen["B2"].Value);
            Assert.Equal("=B1*1.16", resumen["C1"].Formula);
            Assert.Null(resumen["C1"].Value);
            Assert.NotNull(reopened.FindDefinedName("Tabla"));
            Assert.Single(resumen.DataValidations);
            Assert.Equal("D1:E10", resumen.AutoFilter!.Range);
            Assert.Single(resumen.Tables);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void TypedExceptionHandling_CatchesInvalidFile()
    {
        string directory = TestData.NewTemporaryDirectory();
        string path = Path.Combine(directory, "entrada.docx");
        try
        {
            File.WriteAllBytes(path, TestData.NotAZip());

            string message;
            try
            {
                using WordDocument document = WordDocument.Open(path);
                Assert.Fail("Opening a non-Open XML file must throw.");
                return;
            }
            catch (InvalidOfficeFileException ex)
            {
                message = ex.Message;
            }

            Assert.Contains("Office Open XML", message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }
}
