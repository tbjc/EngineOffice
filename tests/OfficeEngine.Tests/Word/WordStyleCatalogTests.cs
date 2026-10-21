using System;
using DocumentFormat.OpenXml;
using OfficeEngine;
using OfficeEngine.Tests.Support;
using OfficeEngine.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeEngine.Tests.Word;

/// <summary>
/// Increment 2.3: the read-only style catalogue (<c>word/styles.xml</c>) and the formatting a style
/// contributes, including the resolution of the <c>w:basedOn</c> inheritance chain.
/// </summary>
public sealed class WordStyleCatalogTests
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void Styles_IsEmpty_WhenTheDocumentHasNoStylePart()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Empty(document.Styles);
        Assert.Null(document.FindStyle("Normal"));
    }

    [Fact]
    public void Styles_IsEmpty_WhenTheStylePartHasNoCatalogue()
    {
        byte[] package = OoxmlValidation.CreateDocxWithEmptyStylePart("body");

        using WordDocument document = WordDocument.Open(TestPackage.ToStream(package));

        Assert.Empty(document.Styles);
        Assert.Null(document.FindStyle("Normal"));
    }

    [Fact]
    public void Styles_ReturnsTheDefinitionsInDocumentOrder()
    {
        using WordDocument document = OpenWithStyles(
            Style("Normal", W.StyleValues.Paragraph, name: "Normal", isDefault: true),
            Style("Heading1", W.StyleValues.Paragraph, name: "Heading 1", basedOn: "Normal"));

        Assert.Equal(2, document.Styles.Count);
        Assert.Equal("Normal", document.Styles[0].Id);
        Assert.Equal("Heading1", document.Styles[1].Id);
    }

    [Fact]
    public void Style_ExposesItsIdentity()
    {
        using WordDocument document = WordDocument.Open(TestPackage.ToStream(
            OoxmlValidation.CreateDocxWithStyles("body", Style(
                "Heading1",
                W.StyleValues.Paragraph,
                name: "Heading 1",
                isDefault: true,
                basedOn: "Normal"))));

        WordStyle style = document.Styles[0];

        Assert.Equal("Heading1", style.Id);
        Assert.Equal("Heading 1", style.Name);
        Assert.Equal(WordStyleType.Paragraph, style.Type);
        Assert.True(style.IsDefault);
        Assert.Equal("Normal", style.BasedOnId);
        Assert.Equal("Heading1", document.FindStyle("Heading1")!.Id);
    }

    [Fact]
    public void Style_Id_IsNull_WhenTheDefinitionHasNoIdentifier()
    {
        W.Style definition = new() { Type = W.StyleValues.Paragraph };
        using WordDocument document = OpenWithStyles(definition);

        Assert.Null(document.Styles[0].Id);
        Assert.Equal("(style without an id)", document.Styles[0].ToString());
    }

    [Fact]
    public void Style_Name_IsNull_WhenTheDefinitionHasNoName()
    {
        using WordDocument document = OpenWithStyles(Style("S", W.StyleValues.Paragraph));

        Assert.Null(document.Styles[0].Name);
    }

    [Fact]
    public void Style_Type_IsNull_WhenTheDefinitionDoesNotDeclareOne()
    {
        W.Style style = new() { StyleId = "S" };

        using WordDocument document = OpenWithStyles(style);

        Assert.Null(document.Styles[0].Type);
    }

    [Theory]
    [InlineData("paragraph", WordStyleType.Paragraph)]
    [InlineData("character", WordStyleType.Character)]
    [InlineData("table", WordStyleType.Table)]
    [InlineData("numbering", WordStyleType.Numbering)]
    public void Style_Type_MapsEveryValueOfficeOpenXmlDefines(string stored, WordStyleType expected)
    {
        W.Style style = new() { StyleId = "S" };
        style.SetAttribute(new OpenXmlAttribute("w", "type", WordprocessingNamespace, stored));

        using WordDocument document = OpenWithStyles(style);

        Assert.Equal(expected, document.Styles[0].Type);
    }

    [Fact]
    public void Style_Type_Throws_WhenTheStoredValueCannotBeParsed()
    {
        W.Style style = new() { StyleId = "S" };
        style.SetAttribute(new OpenXmlAttribute("w", "type", WordprocessingNamespace, "bogus"));
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].Type);
    }

    [Fact]
    public void Style_Name_Throws_WhenTheValueIsMissing()
    {
        W.Style style = new() { StyleId = "S", Type = W.StyleValues.Paragraph };
        style.AppendChild(new W.StyleName());
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].Name);
    }

    [Fact]
    public void Style_BasedOnId_IsNull_WhenTheStyleDoesNotInherit()
    {
        using WordDocument document = OpenWithStyles(Style("Normal", W.StyleValues.Paragraph));

        Assert.Null(document.Styles[0].BasedOnId);
    }

    [Fact]
    public void Style_BasedOnId_Throws_WhenTheValueIsMissing()
    {
        W.Style style = new() { StyleId = "S", Type = W.StyleValues.Paragraph };
        style.AppendChild(new W.BasedOn());
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].BasedOnId);
    }

    [Fact]
    public void Style_IsDefault_IsFalse_WhenTheFlagIsAbsent()
    {
        using WordDocument document = OpenWithStyles(Style("Custom", W.StyleValues.Paragraph));

        Assert.False(document.Styles[0].IsDefault);
    }

    [Fact]
    public void Style_IsDefault_Throws_WhenTheFlagCannotBeParsed()
    {
        W.Style style = new() { StyleId = "S", Type = W.StyleValues.Paragraph };
        style.SetAttribute(new OpenXmlAttribute("w", "default", WordprocessingNamespace, "bogus"));
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].IsDefault);
    }

    [Fact]
    public void FindStyle_ComparesIdentifiersWithOrdinalRules()
    {
        using WordDocument document = OpenWithStyles(Style("Normal", W.StyleValues.Paragraph));

        Assert.NotNull(document.FindStyle("Normal"));
        Assert.Null(document.FindStyle("normal"));
    }

    [Fact]
    public void FindStyle_ReturnsTheFirstDefinitionWithTheIdentifier()
    {
        using WordDocument document = OpenWithStyles(
            Style("S", W.StyleValues.Paragraph, name: "First"),
            Style("S", W.StyleValues.Paragraph, name: "Second"));

        Assert.Equal("First", document.FindStyle("S")?.Name);
    }

    [Fact]
    public void FindStyle_RejectsANullOrEmptyIdentifier()
    {
        using WordDocument document = WordDocument.Create();

        Assert.Throws<ArgumentNullException>(() => document.FindStyle(null!));
        Assert.Throws<ArgumentException>(() => document.FindStyle(string.Empty));
    }

    [Fact]
    public void OwnFormatting_ReadsCharacterProperties()
    {
        W.Style style = Style("Strong", W.StyleValues.Character, runProperties: new W.StyleRunProperties(
            new W.RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" },
            new W.Bold(),
            new W.Italic(),
            new W.Color { Val = "1F3864" },
            new W.FontSize { Val = "28" }));
        using WordDocument document = OpenWithStyles(style);

        WordStyleFormatting formatting = document.Styles[0].OwnFormatting;

        Assert.Equal("Georgia", formatting.FontFamily);
        Assert.Equal(14.0, formatting.FontSizePoints);
        Assert.True(formatting.Bold);
        Assert.True(formatting.Italic);
        Assert.Equal(RgbColor.FromRgb(0x1F, 0x38, 0x64), formatting.Color);
    }

    [Fact]
    public void OwnFormatting_ReadsExplicitlyDisabledBoldAndItalic()
    {
        W.Style style = Style("Plain", W.StyleValues.Character, runProperties: new W.StyleRunProperties(
            new W.Bold { Val = false },
            new W.Italic { Val = false }));
        using WordDocument document = OpenWithStyles(style);

        WordStyleFormatting formatting = document.Styles[0].OwnFormatting;

        Assert.False(formatting.Bold);
        Assert.False(formatting.Italic);
    }

    [Fact]
    public void OwnFormatting_ReadsParagraphProperties()
    {
        W.Style style = Style("Custom", W.StyleValues.Paragraph, paragraphProperties: new W.StyleParagraphProperties(
            new W.SpacingBetweenLines { Before = "120", After = "240" },
            new W.Indentation { Left = "360", Right = "720", FirstLine = "240", Hanging = "120" },
            new W.Justification { Val = W.JustificationValues.Center }));
        using WordDocument document = OpenWithStyles(style);

        WordStyleFormatting formatting = document.Styles[0].OwnFormatting;

        Assert.Equal(6.0, formatting.SpacingBeforePoints);
        Assert.Equal(12.0, formatting.SpacingAfterPoints);
        Assert.Equal(18.0, formatting.LeftIndentPoints);
        Assert.Equal(36.0, formatting.RightIndentPoints);
        Assert.Equal(12.0, formatting.FirstLineIndentPoints);
        Assert.Equal(6.0, formatting.HangingIndentPoints);
        Assert.Equal(ParagraphAlignment.Center, formatting.Alignment);
    }

    [Fact]
    public void OwnFormatting_IsEmpty_WhenTheStyleDefinesNothing()
    {
        using WordDocument document = OpenWithStyles(Style("Normal", W.StyleValues.Paragraph, name: "Normal"));

        Assert.Equal(default(WordStyleFormatting), document.Styles[0].OwnFormatting);
    }

    [Fact]
    public void OwnFormatting_UsesTheHighAnsiFont_WhenTheAsciiFontIsAbsent()
    {
        W.Style style = Style("Custom", W.StyleValues.Character, runProperties: new W.StyleRunProperties(
            new W.RunFonts { HighAnsi = "Cambria" }));
        using WordDocument document = OpenWithStyles(style);

        Assert.Equal("Cambria", document.Styles[0].OwnFormatting.FontFamily);
    }

    [Fact]
    public void OwnFormatting_IgnoresLineSpacing()
    {
        W.Style style = Style("Custom", W.StyleValues.Paragraph, paragraphProperties: new W.StyleParagraphProperties(
            new W.SpacingBetweenLines { Before = "120", Line = "360", LineRule = W.LineSpacingRuleValues.Auto }));
        using WordDocument document = OpenWithStyles(style);

        WordStyleFormatting formatting = document.Styles[0].OwnFormatting;

        Assert.Equal(6.0, formatting.SpacingBeforePoints);
        Assert.Null(formatting.SpacingAfterPoints);
        Assert.Null(formatting.Alignment);
    }

    [Fact]
    public void OwnFormatting_ReportsAThemeColourAsNoFixedColour()
    {
        W.Style style = Style("Accent", W.StyleValues.Character, runProperties: new W.StyleRunProperties(
            new W.Color { Val = "auto", ThemeColor = W.ThemeColorValues.Accent1 }));
        using WordDocument document = OpenWithStyles(style);

        Assert.Null(document.Styles[0].OwnFormatting.Color);
    }

    [Fact]
    public void OwnFormatting_ReportsAColourWithoutAValueAsNoFixedColour()
    {
        W.Style style = Style("Accent", W.StyleValues.Character, runProperties: new W.StyleRunProperties(
            new W.Color { ThemeColor = W.ThemeColorValues.Accent1 }));
        using WordDocument document = OpenWithStyles(style);

        Assert.Null(document.Styles[0].OwnFormatting.Color);
    }

    [Fact]
    public void OwnFormatting_Throws_WhenTheFontSizeCannotBeParsed()
    {
        W.Style style = Style("Broken", W.StyleValues.Character, runProperties: new W.StyleRunProperties(
            new W.FontSize { Val = "not-a-number" }));
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].OwnFormatting);
    }

    [Fact]
    public void OwnFormatting_Throws_WhenAnIndentCannotBeParsed()
    {
        W.Style style = Style("Broken", W.StyleValues.Paragraph, paragraphProperties: new W.StyleParagraphProperties(
            new W.Indentation { Left = "not-a-number" }));
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].OwnFormatting);
    }

    [Fact]
    public void OwnFormatting_Throws_WhenSpacingCannotBeParsed()
    {
        W.Style style = Style("Broken", W.StyleValues.Paragraph, paragraphProperties: new W.StyleParagraphProperties(
            new W.SpacingBetweenLines { Before = "not-a-number" }));
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].OwnFormatting);
    }

    [Fact]
    public void OwnFormatting_Throws_WhenBoldCannotBeParsed()
    {
        W.Bold bold = new();
        bold.SetAttribute(new OpenXmlAttribute("w", "val", WordprocessingNamespace, "bogus"));
        W.Style style = Style("Broken", W.StyleValues.Character, runProperties: new W.StyleRunProperties(bold));
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].OwnFormatting);
    }

    [Theory]
    [InlineData("left", ParagraphAlignment.Left)]
    [InlineData("center", ParagraphAlignment.Center)]
    [InlineData("right", ParagraphAlignment.Right)]
    [InlineData("both", ParagraphAlignment.Both)]
    [InlineData("distribute", ParagraphAlignment.Justify)]
    public void OwnFormatting_MapsEveryModelledAlignment(string stored, ParagraphAlignment expected)
    {
        W.Justification justification = new();
        justification.SetAttribute(new OpenXmlAttribute("w", "val", WordprocessingNamespace, stored));
        W.Style style = Style("Custom", W.StyleValues.Paragraph, paragraphProperties: new W.StyleParagraphProperties(justification));
        using WordDocument document = OpenWithStyles(style);

        Assert.Equal(expected, document.Styles[0].OwnFormatting.Alignment);
    }

    [Fact]
    public void OwnFormatting_ThrowsUnsupported_WhenTheAlignmentIsNotModelled()
    {
        W.Style style = Style("Broken", W.StyleValues.Paragraph, paragraphProperties: new W.StyleParagraphProperties(
            new W.Justification { Val = W.JustificationValues.Start }));
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<UnsupportedOfficeFeatureException>(() => document.Styles[0].OwnFormatting);
    }

    [Fact]
    public void EffectiveFormatting_InheritsFromTheBaseStyle()
    {
        W.Style baseStyle = Style("Base", W.StyleValues.Paragraph, runProperties: new W.StyleRunProperties(
            new W.RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" },
            new W.Bold(),
            new W.FontSize { Val = "24" }));
        W.Style derived = Style("Derived", W.StyleValues.Paragraph, basedOn: "Base");
        using WordDocument document = OpenWithStyles(baseStyle, derived);

        WordStyleFormatting formatting = document.FindStyle("Derived")!.EffectiveFormatting;

        Assert.Equal("Georgia", formatting.FontFamily);
        Assert.True(formatting.Bold);
        Assert.Equal(12.0, formatting.FontSizePoints);
    }

    [Fact]
    public void EffectiveFormatting_AppliesOverridesAlongTheChain()
    {
        W.Style grandparent = Style("Grandparent", W.StyleValues.Paragraph, runProperties: new W.StyleRunProperties(
            new W.RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" },
            new W.Bold(),
            new W.FontSize { Val = "24" }));
        W.Style parent = Style("Parent", W.StyleValues.Paragraph, basedOn: "Grandparent", runProperties: new W.StyleRunProperties(
            new W.Bold { Val = false },
            new W.FontSize { Val = "28" }));
        W.Style child = Style("Child", W.StyleValues.Paragraph, basedOn: "Parent", runProperties: new W.StyleRunProperties(
            new W.Color { Val = "FF0000" }));
        using WordDocument document = OpenWithStyles(grandparent, parent, child);

        WordStyleFormatting formatting = document.FindStyle("Child")!.EffectiveFormatting;

        Assert.Equal("Georgia", formatting.FontFamily);
        Assert.False(formatting.Bold);
        Assert.Equal(14.0, formatting.FontSizePoints);
        Assert.Equal(RgbColor.FromRgb(0xFF, 0x00, 0x00), formatting.Color);
    }

    [Fact]
    public void EffectiveFormatting_KeepsTheInheritedFont_WhenTheDerivedStyleOnlySetsComplexScript()
    {
        W.Style baseStyle = Style("Base", W.StyleValues.Character, runProperties: new W.StyleRunProperties(
            new W.RunFonts { Ascii = "Georgia", HighAnsi = "Georgia" }));
        W.Style derived = Style("Derived", W.StyleValues.Character, basedOn: "Base", runProperties: new W.StyleRunProperties(
            new W.RunFonts { ComplexScript = "Arial" }));
        using WordDocument document = OpenWithStyles(baseStyle, derived);

        Assert.Equal("Georgia", document.FindStyle("Derived")!.EffectiveFormatting.FontFamily);
    }

    [Fact]
    public void EffectiveFormatting_AColourWithoutAFixedRgbOverridesTheInheritedRgbColour()
    {
        W.Style baseStyle = Style("Base", W.StyleValues.Character, runProperties: new W.StyleRunProperties(
            new W.Color { Val = "FF0000" }));
        W.Style derived = Style("Derived", W.StyleValues.Character, basedOn: "Base", runProperties: new W.StyleRunProperties(
            new W.Color { Val = "auto", ThemeColor = W.ThemeColorValues.Accent1 }));
        using WordDocument document = OpenWithStyles(baseStyle, derived);

        Assert.Null(document.FindStyle("Derived")!.EffectiveFormatting.Color);
    }

    [Fact]
    public void EffectiveFormatting_IgnoresADanglingBaseStyle()
    {
        W.Style style = Style("Derived", W.StyleValues.Paragraph, basedOn: "Missing", runProperties: new W.StyleRunProperties(
            new W.Bold()));
        using WordDocument document = OpenWithStyles(style);

        Assert.True(document.Styles[0].EffectiveFormatting.Bold);
    }

    [Fact]
    public void EffectiveFormatting_Throws_WhenTheChainHasACycle()
    {
        W.Style first = Style("A", W.StyleValues.Paragraph, basedOn: "B");
        W.Style second = Style("B", W.StyleValues.Paragraph, basedOn: "A");
        using WordDocument document = OpenWithStyles(first, second);

        Assert.Throws<InvalidOfficeFileException>(() => document.FindStyle("A")!.EffectiveFormatting);
        Assert.Throws<InvalidOfficeFileException>(() => document.FindStyle("B")!.EffectiveFormatting);
    }

    [Fact]
    public void EffectiveFormatting_Throws_WhenAStyleIsBasedOnItself()
    {
        W.Style style = Style("A", W.StyleValues.Paragraph, basedOn: "A");
        using WordDocument document = OpenWithStyles(style);

        Assert.Throws<InvalidOfficeFileException>(() => document.Styles[0].EffectiveFormatting);
    }

    [Fact]
    public void EffectiveFormatting_HandlesAThreeLevelChain()
    {
        W.Style third = Style("C", W.StyleValues.Paragraph, basedOn: "B", runProperties: new W.StyleRunProperties(
            new W.Color { Val = "00FF00" }));
        W.Style second = Style("B", W.StyleValues.Paragraph, basedOn: "A", runProperties: new W.StyleRunProperties(
            new W.FontSize { Val = "24" }));
        W.Style first = Style("A", W.StyleValues.Paragraph, runProperties: new W.StyleRunProperties(
            new W.Bold()));
        using WordDocument document = OpenWithStyles(first, second, third);

        WordStyleFormatting formatting = document.FindStyle("C")!.EffectiveFormatting;

        Assert.True(formatting.Bold);
        Assert.Equal(12.0, formatting.FontSizePoints);
        Assert.Equal(RgbColor.FromRgb(0x00, 0xFF, 0x00), formatting.Color);
    }

    [Fact]
    public void WordStyleFormatting_HasValueEquality()
    {
        W.StyleRunProperties same = new(new W.Bold(), new W.FontSize { Val = "24" });
        using WordDocument document = OpenWithStyles(
            Style("A", W.StyleValues.Character, runProperties: new W.StyleRunProperties(new W.Bold(), new W.FontSize { Val = "24" })),
            Style("B", W.StyleValues.Character, runProperties: same),
            Style("C", W.StyleValues.Character, runProperties: new W.StyleRunProperties(new W.Italic(), new W.FontSize { Val = "24" })));

        WordStyleFormatting first = document.Styles[0].OwnFormatting;
        WordStyleFormatting second = document.Styles[1].OwnFormatting;
        WordStyleFormatting third = document.Styles[2].OwnFormatting;

        Assert.True(first == second);
        Assert.False(first != second);
        Assert.False(first == third);
        Assert.True(first != third);
        Assert.Equal(first.GetHashCode(), second.GetHashCode());
        Assert.True(first.Equals((object)second));
        Assert.False(first.Equals("not a formatting"));
    }

    [Fact]
    public void WordStyle_ToString_ReturnsTheIdentifier()
    {
        using WordDocument document = OpenWithStyles(Style("Normal", W.StyleValues.Paragraph));

        Assert.Equal("Normal", document.Styles[0].ToString());
    }

    [Fact]
    public void WordStyle_ToString_ReportsAStyleWithoutAnIdentifier()
    {
        using WordDocument document = OpenWithStyles(new W.Style { Type = W.StyleValues.Paragraph });

        Assert.Equal("(style without an id)", document.Styles[0].ToString());
    }

    private static WordDocument OpenWithStyles(params W.Style[] styles)
        => WordDocument.Open(TestPackage.ToStream(OoxmlValidation.CreateDocxWithStyles("body", styles)));

    private static W.Style Style(
        string id,
        W.StyleValues type,
        string? name = null,
        bool isDefault = false,
        string? basedOn = null,
        W.StyleParagraphProperties? paragraphProperties = null,
        W.StyleRunProperties? runProperties = null)
    {
        W.Style style = new() { Type = type, StyleId = id };
        if (isDefault)
        {
            style.Default = true;
        }

        if (name is not null)
        {
            style.AppendChild(new W.StyleName { Val = name });
        }

        if (basedOn is not null)
        {
            style.AppendChild(new W.BasedOn { Val = basedOn });
        }

        if (paragraphProperties is not null)
        {
            style.AppendChild(paragraphProperties);
        }

        if (runProperties is not null)
        {
            style.AppendChild(runProperties);
        }

        return style;
    }
}
