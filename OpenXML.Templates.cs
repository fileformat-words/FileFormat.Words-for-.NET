using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using A = DocumentFormat.OpenXml.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using OOXML = OpenXML.Words;
using OOXMLD = OpenXML.Words.Data;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;

namespace OpenXML.Templates
{
    internal class DefaultTemplate
    {
        // Adds child parts and generates content of the specified part.
        public void CreateMainDocumentPart(MainDocumentPart part)
        {
            WebSettingsPart webSettingsPart1 = part.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            DocumentSettingsPart documentSettingsPart1 = part.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = part.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            ThemePart themePart1 = part.AddNewPart<ThemePart>("rId5");
            GenerateThemePart1Content(themePart1);

            FontTablePart fontTablePart1 = part.AddNewPart<FontTablePart>("rId4");
            GenerateFontTablePart1Content(fontTablePart1);

            GeneratePartContent(part);

        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15" } };
            webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            webSettings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
            AllowPNG allowPNG1 = new AllowPNG();

            webSettings1.Append(optimizeForBrowser1);
            webSettings1.Append(allowPNG1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15" } };
            settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            settings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "100" };
            TrackRevisions trackRevisions1 = new TrackRevisions() { Val = false };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 720 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };

            Compatibility compatibility1 = new Compatibility();
            CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" };
            CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
            CompatibilitySetting compatibilitySetting5 = new CompatibilitySetting() { Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };

            compatibility1.Append(compatibilitySetting1);
            compatibility1.Append(compatibilitySetting2);
            compatibility1.Append(compatibilitySetting3);
            compatibility1.Append(compatibilitySetting4);
            compatibility1.Append(compatibilitySetting5);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "7E47A655" };
            Rsid rsid1 = new Rsid() { Val = "0CD935D9" };
            Rsid rsid2 = new Rsid() { Val = "1D0BED6E" };
            Rsid rsid3 = new Rsid() { Val = "7E47A655" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "en-US" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 1026 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };
            W14.DocumentId documentId1 = new W14.DocumentId() { Val = "7E47A655" };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{5ED29589-2441-4447-837E-44C9FF1EE3BD}" };

            settings1.Append(zoom1);
            settings1.Append(trackRevisions1);
            settings1.Append(defaultTabStop1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);
            settings1.Append(documentId1);
            settings1.Append(chartTrackingRefBased1);
            settings1.Append(persistentDocumentId1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            styles1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize1 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "en-US", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(fontSize1);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines1);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 371 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "index 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "index 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "index 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "index 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "index 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "index 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "index 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "index 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "index 9", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "footer", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "index heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "line number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "page number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "macro", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Number", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Closing", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Date", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Normal Table", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "No List", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Table Web 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Table Theme", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Revision", SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "Plain Table 1", UiPriority = 41 };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "Plain Table 2", UiPriority = 42 };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "Plain Table 3", UiPriority = 43 };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "Plain Table 4", UiPriority = 44 };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "Plain Table 5", UiPriority = 45 };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "Grid Table Light", UiPriority = 40 };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "Grid Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "Grid Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() { Name = "Grid Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo() { Name = "List Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo() { Name = "List Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo() { Name = "List Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);
            latentStyles1.Append(latentStyleExceptionInfo138);
            latentStyles1.Append(latentStyleExceptionInfo139);
            latentStyles1.Append(latentStyleExceptionInfo140);
            latentStyles1.Append(latentStyleExceptionInfo141);
            latentStyles1.Append(latentStyleExceptionInfo142);
            latentStyles1.Append(latentStyleExceptionInfo143);
            latentStyles1.Append(latentStyleExceptionInfo144);
            latentStyles1.Append(latentStyleExceptionInfo145);
            latentStyles1.Append(latentStyleExceptionInfo146);
            latentStyles1.Append(latentStyleExceptionInfo147);
            latentStyles1.Append(latentStyleExceptionInfo148);
            latentStyles1.Append(latentStyleExceptionInfo149);
            latentStyles1.Append(latentStyleExceptionInfo150);
            latentStyles1.Append(latentStyleExceptionInfo151);
            latentStyles1.Append(latentStyleExceptionInfo152);
            latentStyles1.Append(latentStyleExceptionInfo153);
            latentStyles1.Append(latentStyleExceptionInfo154);
            latentStyles1.Append(latentStyleExceptionInfo155);
            latentStyles1.Append(latentStyleExceptionInfo156);
            latentStyles1.Append(latentStyleExceptionInfo157);
            latentStyles1.Append(latentStyleExceptionInfo158);
            latentStyles1.Append(latentStyleExceptionInfo159);
            latentStyles1.Append(latentStyleExceptionInfo160);
            latentStyles1.Append(latentStyleExceptionInfo161);
            latentStyles1.Append(latentStyleExceptionInfo162);
            latentStyles1.Append(latentStyleExceptionInfo163);
            latentStyles1.Append(latentStyleExceptionInfo164);
            latentStyles1.Append(latentStyleExceptionInfo165);
            latentStyles1.Append(latentStyleExceptionInfo166);
            latentStyles1.Append(latentStyleExceptionInfo167);
            latentStyles1.Append(latentStyleExceptionInfo168);
            latentStyles1.Append(latentStyleExceptionInfo169);
            latentStyles1.Append(latentStyleExceptionInfo170);
            latentStyles1.Append(latentStyleExceptionInfo171);
            latentStyles1.Append(latentStyleExceptionInfo172);
            latentStyles1.Append(latentStyleExceptionInfo173);
            latentStyles1.Append(latentStyleExceptionInfo174);
            latentStyles1.Append(latentStyleExceptionInfo175);
            latentStyles1.Append(latentStyleExceptionInfo176);
            latentStyles1.Append(latentStyleExceptionInfo177);
            latentStyles1.Append(latentStyleExceptionInfo178);
            latentStyles1.Append(latentStyleExceptionInfo179);
            latentStyles1.Append(latentStyleExceptionInfo180);
            latentStyles1.Append(latentStyleExceptionInfo181);
            latentStyles1.Append(latentStyleExceptionInfo182);
            latentStyles1.Append(latentStyleExceptionInfo183);
            latentStyles1.Append(latentStyleExceptionInfo184);
            latentStyles1.Append(latentStyleExceptionInfo185);
            latentStyles1.Append(latentStyleExceptionInfo186);
            latentStyles1.Append(latentStyleExceptionInfo187);
            latentStyles1.Append(latentStyleExceptionInfo188);
            latentStyles1.Append(latentStyleExceptionInfo189);
            latentStyles1.Append(latentStyleExceptionInfo190);
            latentStyles1.Append(latentStyleExceptionInfo191);
            latentStyles1.Append(latentStyleExceptionInfo192);
            latentStyles1.Append(latentStyleExceptionInfo193);
            latentStyles1.Append(latentStyleExceptionInfo194);
            latentStyles1.Append(latentStyleExceptionInfo195);
            latentStyles1.Append(latentStyleExceptionInfo196);
            latentStyles1.Append(latentStyleExceptionInfo197);
            latentStyles1.Append(latentStyleExceptionInfo198);
            latentStyles1.Append(latentStyleExceptionInfo199);
            latentStyles1.Append(latentStyleExceptionInfo200);
            latentStyles1.Append(latentStyleExceptionInfo201);
            latentStyles1.Append(latentStyleExceptionInfo202);
            latentStyles1.Append(latentStyleExceptionInfo203);
            latentStyles1.Append(latentStyleExceptionInfo204);
            latentStyles1.Append(latentStyleExceptionInfo205);
            latentStyles1.Append(latentStyleExceptionInfo206);
            latentStyles1.Append(latentStyleExceptionInfo207);
            latentStyles1.Append(latentStyleExceptionInfo208);
            latentStyles1.Append(latentStyleExceptionInfo209);
            latentStyles1.Append(latentStyleExceptionInfo210);
            latentStyles1.Append(latentStyleExceptionInfo211);
            latentStyles1.Append(latentStyleExceptionInfo212);
            latentStyles1.Append(latentStyleExceptionInfo213);
            latentStyles1.Append(latentStyleExceptionInfo214);
            latentStyles1.Append(latentStyleExceptionInfo215);
            latentStyles1.Append(latentStyleExceptionInfo216);
            latentStyles1.Append(latentStyleExceptionInfo217);
            latentStyles1.Append(latentStyleExceptionInfo218);
            latentStyles1.Append(latentStyleExceptionInfo219);
            latentStyles1.Append(latentStyleExceptionInfo220);
            latentStyles1.Append(latentStyleExceptionInfo221);
            latentStyles1.Append(latentStyleExceptionInfo222);
            latentStyles1.Append(latentStyleExceptionInfo223);
            latentStyles1.Append(latentStyleExceptionInfo224);
            latentStyles1.Append(latentStyleExceptionInfo225);
            latentStyles1.Append(latentStyleExceptionInfo226);
            latentStyles1.Append(latentStyleExceptionInfo227);
            latentStyles1.Append(latentStyleExceptionInfo228);
            latentStyles1.Append(latentStyleExceptionInfo229);
            latentStyles1.Append(latentStyleExceptionInfo230);
            latentStyles1.Append(latentStyleExceptionInfo231);
            latentStyles1.Append(latentStyleExceptionInfo232);
            latentStyles1.Append(latentStyleExceptionInfo233);
            latentStyles1.Append(latentStyleExceptionInfo234);
            latentStyles1.Append(latentStyleExceptionInfo235);
            latentStyles1.Append(latentStyleExceptionInfo236);
            latentStyles1.Append(latentStyleExceptionInfo237);
            latentStyles1.Append(latentStyleExceptionInfo238);
            latentStyles1.Append(latentStyleExceptionInfo239);
            latentStyles1.Append(latentStyleExceptionInfo240);
            latentStyles1.Append(latentStyleExceptionInfo241);
            latentStyles1.Append(latentStyleExceptionInfo242);
            latentStyles1.Append(latentStyleExceptionInfo243);
            latentStyles1.Append(latentStyleExceptionInfo244);
            latentStyles1.Append(latentStyleExceptionInfo245);
            latentStyles1.Append(latentStyleExceptionInfo246);
            latentStyles1.Append(latentStyleExceptionInfo247);
            latentStyles1.Append(latentStyleExceptionInfo248);
            latentStyles1.Append(latentStyleExceptionInfo249);
            latentStyles1.Append(latentStyleExceptionInfo250);
            latentStyles1.Append(latentStyleExceptionInfo251);
            latentStyles1.Append(latentStyleExceptionInfo252);
            latentStyles1.Append(latentStyleExceptionInfo253);
            latentStyles1.Append(latentStyleExceptionInfo254);
            latentStyles1.Append(latentStyleExceptionInfo255);
            latentStyles1.Append(latentStyleExceptionInfo256);
            latentStyles1.Append(latentStyleExceptionInfo257);
            latentStyles1.Append(latentStyleExceptionInfo258);
            latentStyles1.Append(latentStyleExceptionInfo259);
            latentStyles1.Append(latentStyleExceptionInfo260);
            latentStyles1.Append(latentStyleExceptionInfo261);
            latentStyles1.Append(latentStyleExceptionInfo262);
            latentStyles1.Append(latentStyleExceptionInfo263);
            latentStyles1.Append(latentStyleExceptionInfo264);
            latentStyles1.Append(latentStyleExceptionInfo265);
            latentStyles1.Append(latentStyleExceptionInfo266);
            latentStyles1.Append(latentStyleExceptionInfo267);
            latentStyles1.Append(latentStyleExceptionInfo268);
            latentStyles1.Append(latentStyleExceptionInfo269);
            latentStyles1.Append(latentStyleExceptionInfo270);
            latentStyles1.Append(latentStyleExceptionInfo271);
            latentStyles1.Append(latentStyleExceptionInfo272);
            latentStyles1.Append(latentStyleExceptionInfo273);
            latentStyles1.Append(latentStyleExceptionInfo274);
            latentStyles1.Append(latentStyleExceptionInfo275);
            latentStyles1.Append(latentStyleExceptionInfo276);
            latentStyles1.Append(latentStyleExceptionInfo277);
            latentStyles1.Append(latentStyleExceptionInfo278);
            latentStyles1.Append(latentStyleExceptionInfo279);
            latentStyles1.Append(latentStyleExceptionInfo280);
            latentStyles1.Append(latentStyleExceptionInfo281);
            latentStyles1.Append(latentStyleExceptionInfo282);
            latentStyles1.Append(latentStyleExceptionInfo283);
            latentStyles1.Append(latentStyleExceptionInfo284);
            latentStyles1.Append(latentStyleExceptionInfo285);
            latentStyles1.Append(latentStyleExceptionInfo286);
            latentStyles1.Append(latentStyleExceptionInfo287);
            latentStyles1.Append(latentStyleExceptionInfo288);
            latentStyles1.Append(latentStyleExceptionInfo289);
            latentStyles1.Append(latentStyleExceptionInfo290);
            latentStyles1.Append(latentStyleExceptionInfo291);
            latentStyles1.Append(latentStyleExceptionInfo292);
            latentStyles1.Append(latentStyleExceptionInfo293);
            latentStyles1.Append(latentStyleExceptionInfo294);
            latentStyles1.Append(latentStyleExceptionInfo295);
            latentStyles1.Append(latentStyleExceptionInfo296);
            latentStyles1.Append(latentStyleExceptionInfo297);
            latentStyles1.Append(latentStyleExceptionInfo298);
            latentStyles1.Append(latentStyleExceptionInfo299);
            latentStyles1.Append(latentStyleExceptionInfo300);
            latentStyles1.Append(latentStyleExceptionInfo301);
            latentStyles1.Append(latentStyleExceptionInfo302);
            latentStyles1.Append(latentStyleExceptionInfo303);
            latentStyles1.Append(latentStyleExceptionInfo304);
            latentStyles1.Append(latentStyleExceptionInfo305);
            latentStyles1.Append(latentStyleExceptionInfo306);
            latentStyles1.Append(latentStyleExceptionInfo307);
            latentStyles1.Append(latentStyleExceptionInfo308);
            latentStyles1.Append(latentStyleExceptionInfo309);
            latentStyles1.Append(latentStyleExceptionInfo310);
            latentStyles1.Append(latentStyleExceptionInfo311);
            latentStyles1.Append(latentStyleExceptionInfo312);
            latentStyles1.Append(latentStyleExceptionInfo313);
            latentStyles1.Append(latentStyleExceptionInfo314);
            latentStyles1.Append(latentStyleExceptionInfo315);
            latentStyles1.Append(latentStyleExceptionInfo316);
            latentStyles1.Append(latentStyleExceptionInfo317);
            latentStyles1.Append(latentStyleExceptionInfo318);
            latentStyles1.Append(latentStyleExceptionInfo319);
            latentStyles1.Append(latentStyleExceptionInfo320);
            latentStyles1.Append(latentStyleExceptionInfo321);
            latentStyles1.Append(latentStyleExceptionInfo322);
            latentStyles1.Append(latentStyleExceptionInfo323);
            latentStyles1.Append(latentStyleExceptionInfo324);
            latentStyles1.Append(latentStyleExceptionInfo325);
            latentStyles1.Append(latentStyleExceptionInfo326);
            latentStyles1.Append(latentStyleExceptionInfo327);
            latentStyles1.Append(latentStyleExceptionInfo328);
            latentStyles1.Append(latentStyleExceptionInfo329);
            latentStyles1.Append(latentStyleExceptionInfo330);
            latentStyles1.Append(latentStyleExceptionInfo331);
            latentStyles1.Append(latentStyleExceptionInfo332);
            latentStyles1.Append(latentStyleExceptionInfo333);
            latentStyles1.Append(latentStyleExceptionInfo334);
            latentStyles1.Append(latentStyleExceptionInfo335);
            latentStyles1.Append(latentStyleExceptionInfo336);
            latentStyles1.Append(latentStyleExceptionInfo337);
            latentStyles1.Append(latentStyleExceptionInfo338);
            latentStyles1.Append(latentStyleExceptionInfo339);
            latentStyles1.Append(latentStyleExceptionInfo340);
            latentStyles1.Append(latentStyleExceptionInfo341);
            latentStyles1.Append(latentStyleExceptionInfo342);
            latentStyles1.Append(latentStyleExceptionInfo343);
            latentStyles1.Append(latentStyleExceptionInfo344);
            latentStyles1.Append(latentStyleExceptionInfo345);
            latentStyles1.Append(latentStyleExceptionInfo346);
            latentStyles1.Append(latentStyleExceptionInfo347);
            latentStyles1.Append(latentStyleExceptionInfo348);
            latentStyles1.Append(latentStyleExceptionInfo349);
            latentStyles1.Append(latentStyleExceptionInfo350);
            latentStyles1.Append(latentStyleExceptionInfo351);
            latentStyles1.Append(latentStyleExceptionInfo352);
            latentStyles1.Append(latentStyleExceptionInfo353);
            latentStyles1.Append(latentStyleExceptionInfo354);
            latentStyles1.Append(latentStyleExceptionInfo355);
            latentStyles1.Append(latentStyleExceptionInfo356);
            latentStyles1.Append(latentStyleExceptionInfo357);
            latentStyles1.Append(latentStyleExceptionInfo358);
            latentStyles1.Append(latentStyleExceptionInfo359);
            latentStyles1.Append(latentStyleExceptionInfo360);
            latentStyles1.Append(latentStyleExceptionInfo361);
            latentStyles1.Append(latentStyleExceptionInfo362);
            latentStyles1.Append(latentStyleExceptionInfo363);
            latentStyles1.Append(latentStyleExceptionInfo364);
            latentStyles1.Append(latentStyleExceptionInfo365);
            latentStyles1.Append(latentStyleExceptionInfo366);
            latentStyles1.Append(latentStyleExceptionInfo367);
            latentStyles1.Append(latentStyleExceptionInfo368);
            latentStyles1.Append(latentStyleExceptionInfo369);
            latentStyles1.Append(latentStyleExceptionInfo370);
            latentStyles1.Append(latentStyleExceptionInfo371);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();

            style1.Append(styleName1);
            style1.Append(primaryStyle1);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "DefaultParagraphFont", Default = true };
            StyleName styleName2 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style2.Append(styleName2);
            style2.Append(uIPriority1);
            style2.Append(semiHidden1);
            style2.Append(unhideWhenUsed1);

            Style style3 = new Style() { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style3.Append(styleName3);
            style3.Append(uIPriority2);
            style3.Append(semiHidden2);
            style3.Append(unhideWhenUsed2);
            style3.Append(styleTableProperties1);

            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "NoList", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed3);

            Style style5 = new Style() { Type = StyleValues.Character, StyleId = "Heading1Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style5.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style5.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName5 = new StyleName() { Val = "Heading 1 Char" };
            styleName5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn1 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "Heading1" };
            linkedStyle1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority4 = new UIPriority() { Val = 9 };
            uIPriority4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            styleRunProperties1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts2 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color1 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize2 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties1.Append(runFonts2);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(fontSize2);
            styleRunProperties1.Append(fontSizeComplexScript2);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(linkedStyle1);
            style5.Append(uIPriority4);
            style5.Append(styleRunProperties1);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style6.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style6.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName6 = new StyleName() { Val = "heading 1" };
            styleName6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };
            basedOn2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "Heading1Char" };
            linkedStyle2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority5 = new UIPriority() { Val = 9 };
            uIPriority5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle2 = new PrimaryStyle();
            primaryStyle2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            styleParagraphProperties1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext1 = new KeepNext();
            keepNext1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines1 = new KeepLines();
            keepLines1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "240", After = "0" };
            spacingBetweenLines2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel1 = new OutlineLevel() { Val = 0 };
            outlineLevel1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties1.Append(keepNext1);
            styleParagraphProperties1.Append(keepLines1);
            styleParagraphProperties1.Append(spacingBetweenLines2);
            styleParagraphProperties1.Append(outlineLevel1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            styleRunProperties2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts3 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color2 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize3 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "32" };

            styleRunProperties2.Append(runFonts3);
            styleRunProperties2.Append(color2);
            styleRunProperties2.Append(fontSize3);
            styleRunProperties2.Append(fontSizeComplexScript3);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(nextParagraphStyle1);
            style6.Append(linkedStyle2);
            style6.Append(uIPriority5);
            style6.Append(primaryStyle2);
            style6.Append(styleParagraphProperties1);
            style6.Append(styleRunProperties2);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading2", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style7.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style7.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName7 = new StyleName() { Val = "heading 2" };
            styleName7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn3 = new BasedOn() { Val = "Normal" };
            basedOn3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "Heading2Char" };
            linkedStyle3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority6 = new UIPriority() { Val = 9 };
            uIPriority6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            unhideWhenUsed4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle3 = new PrimaryStyle();
            primaryStyle3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            styleParagraphProperties2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext2 = new KeepNext();
            keepNext2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines2 = new KeepLines();
            keepLines2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Before = "40", After = "0" };
            spacingBetweenLines3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel2 = new OutlineLevel() { Val = 1 };
            outlineLevel2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties2.Append(keepNext2);
            styleParagraphProperties2.Append(keepLines2);
            styleParagraphProperties2.Append(spacingBetweenLines3);
            styleParagraphProperties2.Append(outlineLevel2);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            styleRunProperties3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts4 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color3 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize4 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties3.Append(runFonts4);
            styleRunProperties3.Append(color3);
            styleRunProperties3.Append(fontSize4);
            styleRunProperties3.Append(fontSizeComplexScript4);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(nextParagraphStyle2);
            style7.Append(linkedStyle3);
            style7.Append(uIPriority6);
            style7.Append(unhideWhenUsed4);
            style7.Append(primaryStyle3);
            style7.Append(styleParagraphProperties2);
            style7.Append(styleRunProperties3);

            Style style8 = new Style() { Type = StyleValues.Character, StyleId = "Heading2Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style8.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style8.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName8 = new StyleName() { Val = "Heading 2 Char" };
            styleName8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn4 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "Heading2" };
            linkedStyle4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority7 = new UIPriority() { Val = 9 };
            uIPriority7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            styleRunProperties4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts5 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color4 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };
            FontSize fontSize5 = new FontSize() { Val = "26" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "26" };

            styleRunProperties4.Append(runFonts5);
            styleRunProperties4.Append(color4);
            styleRunProperties4.Append(fontSize5);
            styleRunProperties4.Append(fontSizeComplexScript5);

            style8.Append(styleName8);
            style8.Append(basedOn4);
            style8.Append(linkedStyle4);
            style8.Append(uIPriority7);
            style8.Append(styleRunProperties4);

            Style style9 = new Style() { Type = StyleValues.Character, StyleId = "Heading3Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style9.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style9.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName9 = new StyleName() { Val = "Heading 3 Char" };
            styleName9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn5 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "Heading3" };
            linkedStyle5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority8 = new UIPriority() { Val = 9 };
            uIPriority8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            styleRunProperties5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts6 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color5 = new Color() { Val = "1F4D78", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };
            FontSize fontSize6 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties5.Append(runFonts6);
            styleRunProperties5.Append(color5);
            styleRunProperties5.Append(fontSize6);
            styleRunProperties5.Append(fontSizeComplexScript6);

            style9.Append(styleName9);
            style9.Append(basedOn5);
            style9.Append(linkedStyle5);
            style9.Append(uIPriority8);
            style9.Append(styleRunProperties5);

            Style style10 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading3", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style10.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style10.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName10 = new StyleName() { Val = "heading 3" };
            styleName10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn6 = new BasedOn() { Val = "Normal" };
            basedOn6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle3 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "Heading3Char" };
            linkedStyle6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority9 = new UIPriority() { Val = 9 };
            uIPriority9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
            unhideWhenUsed5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle4 = new PrimaryStyle();
            primaryStyle4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();
            styleParagraphProperties3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext3 = new KeepNext();
            keepNext3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines3 = new KeepLines();
            keepLines3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Before = "40", After = "0" };
            spacingBetweenLines4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel3 = new OutlineLevel() { Val = 2 };
            outlineLevel3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties3.Append(keepNext3);
            styleParagraphProperties3.Append(keepLines3);
            styleParagraphProperties3.Append(spacingBetweenLines4);
            styleParagraphProperties3.Append(outlineLevel3);

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            styleRunProperties6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts7 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color6 = new Color() { Val = "1F4D78", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };
            FontSize fontSize7 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties6.Append(runFonts7);
            styleRunProperties6.Append(color6);
            styleRunProperties6.Append(fontSize7);
            styleRunProperties6.Append(fontSizeComplexScript7);

            style10.Append(styleName10);
            style10.Append(basedOn6);
            style10.Append(nextParagraphStyle3);
            style10.Append(linkedStyle6);
            style10.Append(uIPriority9);
            style10.Append(unhideWhenUsed5);
            style10.Append(primaryStyle4);
            style10.Append(styleParagraphProperties3);
            style10.Append(styleRunProperties6);

            Style style11 = new Style() { Type = StyleValues.Character, StyleId = "Heading4Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style11.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style11.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style11.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName11 = new StyleName() { Val = "Heading 4 Char" };
            styleName11.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn7 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "Heading4" };
            linkedStyle7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority10 = new UIPriority() { Val = 9 };
            uIPriority10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            styleRunProperties7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts8 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            Color color7 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties7.Append(runFonts8);
            styleRunProperties7.Append(italic1);
            styleRunProperties7.Append(italicComplexScript1);
            styleRunProperties7.Append(color7);

            style11.Append(styleName11);
            style11.Append(basedOn7);
            style11.Append(linkedStyle7);
            style11.Append(uIPriority10);
            style11.Append(styleRunProperties7);

            Style style12 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading4", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style12.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style12.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style12.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName12 = new StyleName() { Val = "heading 4" };
            styleName12.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn8 = new BasedOn() { Val = "Normal" };
            basedOn8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle4 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "Heading4Char" };
            linkedStyle8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority11 = new UIPriority() { Val = 9 };
            uIPriority11.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UnhideWhenUsed unhideWhenUsed6 = new UnhideWhenUsed();
            unhideWhenUsed6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle5 = new PrimaryStyle();
            primaryStyle5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            styleParagraphProperties4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext4 = new KeepNext();
            keepNext4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines4 = new KeepLines();
            keepLines4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Before = "40", After = "0" };
            spacingBetweenLines5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel4 = new OutlineLevel() { Val = 3 };
            outlineLevel4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties4.Append(keepNext4);
            styleParagraphProperties4.Append(keepLines4);
            styleParagraphProperties4.Append(spacingBetweenLines5);
            styleParagraphProperties4.Append(outlineLevel4);

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            styleRunProperties8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts9 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic2 = new Italic();
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();
            Color color8 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties8.Append(runFonts9);
            styleRunProperties8.Append(italic2);
            styleRunProperties8.Append(italicComplexScript2);
            styleRunProperties8.Append(color8);

            style12.Append(styleName12);
            style12.Append(basedOn8);
            style12.Append(nextParagraphStyle4);
            style12.Append(linkedStyle8);
            style12.Append(uIPriority11);
            style12.Append(unhideWhenUsed6);
            style12.Append(primaryStyle5);
            style12.Append(styleParagraphProperties4);
            style12.Append(styleRunProperties8);

            Style style13 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading5", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style13.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style13.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style13.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName13 = new StyleName() { Val = "heading 5" };
            styleName13.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn9 = new BasedOn() { Val = "Normal" };
            basedOn9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle5 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "Heading5Char" };
            linkedStyle9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority12 = new UIPriority() { Val = 9 };
            uIPriority12.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UnhideWhenUsed unhideWhenUsed7 = new UnhideWhenUsed();
            unhideWhenUsed7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle6 = new PrimaryStyle();
            primaryStyle6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
            styleParagraphProperties5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext5 = new KeepNext();
            keepNext5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines5 = new KeepLines();
            keepLines5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Before = "40", After = "0" };
            spacingBetweenLines6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel5 = new OutlineLevel() { Val = 4 };
            outlineLevel5.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties5.Append(keepNext5);
            styleParagraphProperties5.Append(keepLines5);
            styleParagraphProperties5.Append(spacingBetweenLines6);
            styleParagraphProperties5.Append(outlineLevel5);

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            styleRunProperties9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts10 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color9 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties9.Append(runFonts10);
            styleRunProperties9.Append(color9);

            style13.Append(styleName13);
            style13.Append(basedOn9);
            style13.Append(nextParagraphStyle5);
            style13.Append(linkedStyle9);
            style13.Append(uIPriority12);
            style13.Append(unhideWhenUsed7);
            style13.Append(primaryStyle6);
            style13.Append(styleParagraphProperties5);
            style13.Append(styleRunProperties9);

            Style style14 = new Style() { Type = StyleValues.Character, StyleId = "Heading5Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style14.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style14.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style14.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName14 = new StyleName() { Val = "Heading 5 Char" };
            styleName14.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn10 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "Heading5" };
            linkedStyle10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority13 = new UIPriority() { Val = 9 };
            uIPriority13.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            styleRunProperties10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts11 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color10 = new Color() { Val = "2E74B5", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "BF" };

            styleRunProperties10.Append(runFonts11);
            styleRunProperties10.Append(color10);

            style14.Append(styleName14);
            style14.Append(basedOn10);
            style14.Append(linkedStyle10);
            style14.Append(uIPriority13);
            style14.Append(styleRunProperties10);

            Style style15 = new Style() { Type = StyleValues.Character, StyleId = "Heading6Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style15.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style15.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style15.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName15 = new StyleName() { Val = "Heading 6 Char" };
            styleName15.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn11 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn11.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle11 = new LinkedStyle() { Val = "Heading6" };
            linkedStyle11.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority14 = new UIPriority() { Val = 9 };
            uIPriority14.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            styleRunProperties11.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts12 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color11 = new Color() { Val = "1F4D78", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };

            styleRunProperties11.Append(runFonts12);
            styleRunProperties11.Append(color11);

            style15.Append(styleName15);
            style15.Append(basedOn11);
            style15.Append(linkedStyle11);
            style15.Append(uIPriority14);
            style15.Append(styleRunProperties11);

            Style style16 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading6", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style16.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style16.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style16.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName16 = new StyleName() { Val = "heading 6" };
            styleName16.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn12 = new BasedOn() { Val = "Normal" };
            basedOn12.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle6 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle12 = new LinkedStyle() { Val = "Heading6Char" };
            linkedStyle12.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority15 = new UIPriority() { Val = 9 };
            uIPriority15.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UnhideWhenUsed unhideWhenUsed8 = new UnhideWhenUsed();
            unhideWhenUsed8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle7 = new PrimaryStyle();
            primaryStyle7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties6 = new StyleParagraphProperties();
            styleParagraphProperties6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext6 = new KeepNext();
            keepNext6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines6 = new KeepLines();
            keepLines6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Before = "40", After = "0" };
            spacingBetweenLines7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel6 = new OutlineLevel() { Val = 5 };
            outlineLevel6.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties6.Append(keepNext6);
            styleParagraphProperties6.Append(keepLines6);
            styleParagraphProperties6.Append(spacingBetweenLines7);
            styleParagraphProperties6.Append(outlineLevel6);

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            styleRunProperties12.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts13 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color12 = new Color() { Val = "1F4D78", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };

            styleRunProperties12.Append(runFonts13);
            styleRunProperties12.Append(color12);

            style16.Append(styleName16);
            style16.Append(basedOn12);
            style16.Append(nextParagraphStyle6);
            style16.Append(linkedStyle12);
            style16.Append(uIPriority15);
            style16.Append(unhideWhenUsed8);
            style16.Append(primaryStyle7);
            style16.Append(styleParagraphProperties6);
            style16.Append(styleRunProperties12);

            Style style17 = new Style() { Type = StyleValues.Character, StyleId = "Heading7Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style17.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style17.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style17.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName17 = new StyleName() { Val = "Heading 7 Char" };
            styleName17.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn13 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn13.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle13 = new LinkedStyle() { Val = "Heading7" };
            linkedStyle13.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority16 = new UIPriority() { Val = 9 };
            uIPriority16.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            styleRunProperties13.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts14 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic3 = new Italic();
            ItalicComplexScript italicComplexScript3 = new ItalicComplexScript();
            Color color13 = new Color() { Val = "1F4D78", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };

            styleRunProperties13.Append(runFonts14);
            styleRunProperties13.Append(italic3);
            styleRunProperties13.Append(italicComplexScript3);
            styleRunProperties13.Append(color13);

            style17.Append(styleName17);
            style17.Append(basedOn13);
            style17.Append(linkedStyle13);
            style17.Append(uIPriority16);
            style17.Append(styleRunProperties13);

            Style style18 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading7", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style18.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style18.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style18.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName18 = new StyleName() { Val = "heading 7" };
            styleName18.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn14 = new BasedOn() { Val = "Normal" };
            basedOn14.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle7 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle14 = new LinkedStyle() { Val = "Heading7Char" };
            linkedStyle14.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority17 = new UIPriority() { Val = 9 };
            uIPriority17.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UnhideWhenUsed unhideWhenUsed9 = new UnhideWhenUsed();
            unhideWhenUsed9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle8 = new PrimaryStyle();
            primaryStyle8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties7 = new StyleParagraphProperties();
            styleParagraphProperties7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext7 = new KeepNext();
            keepNext7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines7 = new KeepLines();
            keepLines7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Before = "40", After = "0" };
            spacingBetweenLines8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel7 = new OutlineLevel() { Val = 6 };
            outlineLevel7.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties7.Append(keepNext7);
            styleParagraphProperties7.Append(keepLines7);
            styleParagraphProperties7.Append(spacingBetweenLines8);
            styleParagraphProperties7.Append(outlineLevel7);

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            styleRunProperties14.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts15 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic4 = new Italic();
            ItalicComplexScript italicComplexScript4 = new ItalicComplexScript();
            Color color14 = new Color() { Val = "1F4D78", ThemeColor = ThemeColorValues.Accent1, ThemeShade = "7F" };

            styleRunProperties14.Append(runFonts15);
            styleRunProperties14.Append(italic4);
            styleRunProperties14.Append(italicComplexScript4);
            styleRunProperties14.Append(color14);

            style18.Append(styleName18);
            style18.Append(basedOn14);
            style18.Append(nextParagraphStyle7);
            style18.Append(linkedStyle14);
            style18.Append(uIPriority17);
            style18.Append(unhideWhenUsed9);
            style18.Append(primaryStyle8);
            style18.Append(styleParagraphProperties7);
            style18.Append(styleRunProperties14);

            Style style19 = new Style() { Type = StyleValues.Character, StyleId = "Heading8Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style19.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style19.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style19.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName19 = new StyleName() { Val = "Heading 8 Char" };
            styleName19.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn15 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn15.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle15 = new LinkedStyle() { Val = "Heading8" };
            linkedStyle15.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority18 = new UIPriority() { Val = 9 };
            uIPriority18.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties15 = new StyleRunProperties();
            styleRunProperties15.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts16 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color15 = new Color() { Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };
            FontSize fontSize8 = new FontSize() { Val = "21" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "21" };

            styleRunProperties15.Append(runFonts16);
            styleRunProperties15.Append(color15);
            styleRunProperties15.Append(fontSize8);
            styleRunProperties15.Append(fontSizeComplexScript8);

            style19.Append(styleName19);
            style19.Append(basedOn15);
            style19.Append(linkedStyle15);
            style19.Append(uIPriority18);
            style19.Append(styleRunProperties15);

            Style style20 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading8", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style20.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style20.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style20.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName20 = new StyleName() { Val = "heading 8" };
            styleName20.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn16 = new BasedOn() { Val = "Normal" };
            basedOn16.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle8 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle16 = new LinkedStyle() { Val = "Heading8Char" };
            linkedStyle16.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority19 = new UIPriority() { Val = 9 };
            uIPriority19.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UnhideWhenUsed unhideWhenUsed10 = new UnhideWhenUsed();
            unhideWhenUsed10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle9 = new PrimaryStyle();
            primaryStyle9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties8 = new StyleParagraphProperties();
            styleParagraphProperties8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext8 = new KeepNext();
            keepNext8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines8 = new KeepLines();
            keepLines8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Before = "40", After = "0" };
            spacingBetweenLines9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel8 = new OutlineLevel() { Val = 7 };
            outlineLevel8.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties8.Append(keepNext8);
            styleParagraphProperties8.Append(keepLines8);
            styleParagraphProperties8.Append(spacingBetweenLines9);
            styleParagraphProperties8.Append(outlineLevel8);

            StyleRunProperties styleRunProperties16 = new StyleRunProperties();
            styleRunProperties16.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts17 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Color color16 = new Color() { Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };
            FontSize fontSize9 = new FontSize() { Val = "21" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "21" };

            styleRunProperties16.Append(runFonts17);
            styleRunProperties16.Append(color16);
            styleRunProperties16.Append(fontSize9);
            styleRunProperties16.Append(fontSizeComplexScript9);

            style20.Append(styleName20);
            style20.Append(basedOn16);
            style20.Append(nextParagraphStyle8);
            style20.Append(linkedStyle16);
            style20.Append(uIPriority19);
            style20.Append(unhideWhenUsed10);
            style20.Append(primaryStyle9);
            style20.Append(styleParagraphProperties8);
            style20.Append(styleRunProperties16);

            Style style21 = new Style() { Type = StyleValues.Character, StyleId = "Heading9Char", CustomStyle = true, MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style21.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style21.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style21.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName21 = new StyleName() { Val = "Heading 9 Char" };
            styleName21.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn17 = new BasedOn() { Val = "DefaultParagraphFont" };
            basedOn17.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle17 = new LinkedStyle() { Val = "Heading9" };
            linkedStyle17.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority20 = new UIPriority() { Val = 9 };
            uIPriority20.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleRunProperties styleRunProperties17 = new StyleRunProperties();
            styleRunProperties17.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts18 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic5 = new Italic();
            ItalicComplexScript italicComplexScript5 = new ItalicComplexScript();
            Color color17 = new Color() { Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };
            FontSize fontSize10 = new FontSize() { Val = "21" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "21" };

            styleRunProperties17.Append(runFonts18);
            styleRunProperties17.Append(italic5);
            styleRunProperties17.Append(italicComplexScript5);
            styleRunProperties17.Append(color17);
            styleRunProperties17.Append(fontSize10);
            styleRunProperties17.Append(fontSizeComplexScript10);

            style21.Append(styleName21);
            style21.Append(basedOn17);
            style21.Append(linkedStyle17);
            style21.Append(uIPriority20);
            style21.Append(styleRunProperties17);

            Style style22 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading9", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style22.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style22.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style22.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName22 = new StyleName() { Val = "heading 9" };
            styleName22.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn18 = new BasedOn() { Val = "Normal" };
            basedOn18.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            NextParagraphStyle nextParagraphStyle9 = new NextParagraphStyle() { Val = "Normal" };
            nextParagraphStyle9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            LinkedStyle linkedStyle18 = new LinkedStyle() { Val = "Heading9Char" };
            linkedStyle18.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority21 = new UIPriority() { Val = 9 };
            uIPriority21.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UnhideWhenUsed unhideWhenUsed11 = new UnhideWhenUsed();
            unhideWhenUsed11.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            PrimaryStyle primaryStyle10 = new PrimaryStyle();
            primaryStyle10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties9 = new StyleParagraphProperties();
            styleParagraphProperties9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepNext keepNext9 = new KeepNext();
            keepNext9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            KeepLines keepLines9 = new KeepLines();
            keepLines9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Before = "40", After = "0" };
            spacingBetweenLines10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            OutlineLevel outlineLevel9 = new OutlineLevel() { Val = 8 };
            outlineLevel9.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties9.Append(keepNext9);
            styleParagraphProperties9.Append(keepLines9);
            styleParagraphProperties9.Append(spacingBetweenLines10);
            styleParagraphProperties9.Append(outlineLevel9);

            StyleRunProperties styleRunProperties18 = new StyleRunProperties();
            styleRunProperties18.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            RunFonts runFonts19 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
            Italic italic6 = new Italic();
            ItalicComplexScript italicComplexScript6 = new ItalicComplexScript();
            Color color18 = new Color() { Val = "272727", ThemeColor = ThemeColorValues.Text1, ThemeTint = "D8" };
            FontSize fontSize11 = new FontSize() { Val = "21" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "21" };

            styleRunProperties18.Append(runFonts19);
            styleRunProperties18.Append(italic6);
            styleRunProperties18.Append(italicComplexScript6);
            styleRunProperties18.Append(color18);
            styleRunProperties18.Append(fontSize11);
            styleRunProperties18.Append(fontSizeComplexScript11);

            style22.Append(styleName22);
            style22.Append(basedOn18);
            style22.Append(nextParagraphStyle9);
            style22.Append(linkedStyle18);
            style22.Append(uIPriority21);
            style22.Append(unhideWhenUsed11);
            style22.Append(primaryStyle10);
            style22.Append(styleParagraphProperties9);
            style22.Append(styleRunProperties18);

            Style style23 = new Style() { Type = StyleValues.Table, StyleId = "TableGrid" };
            style23.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName23 = new StyleName() { Val = "Table Grid" };
            styleName23.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn19 = new BasedOn() { Val = "TableNormal" };
            basedOn19.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority22 = new UIPriority() { Val = 59 };
            uIPriority22.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Rsid rsid4 = new Rsid() { Val = "00FB4123" };
            rsid4.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties10 = new StyleParagraphProperties();
            styleParagraphProperties10.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties10.Append(spacingBetweenLines11);

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();
            styleTableProperties2.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            TableIndentation tableIndentation2 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", ThemeColor = ThemeColorValues.Text1, Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(topMargin2);
            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(bottomMargin2);
            tableCellMarginDefault2.Append(tableCellRightMargin2);

            styleTableProperties2.Append(tableIndentation2);
            styleTableProperties2.Append(tableBorders1);
            styleTableProperties2.Append(tableCellMarginDefault2);

            style23.Append(styleName23);
            style23.Append(basedOn19);
            style23.Append(uIPriority22);
            style23.Append(rsid4);
            style23.Append(styleParagraphProperties10);
            style23.Append(styleTableProperties2);

            Style style24 = new Style() { Type = StyleValues.Table, StyleId = "TableGridLight", MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            style24.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            style24.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            style24.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleName styleName24 = new StyleName() { Val = "Grid Table Light" };
            styleName24.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            BasedOn basedOn20 = new BasedOn() { Val = "TableNormal" };
            basedOn20.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            UIPriority uIPriority23 = new UIPriority() { Val = 40 };
            uIPriority23.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            StyleParagraphProperties styleParagraphProperties11 = new StyleParagraphProperties();
            styleParagraphProperties11.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };
            spacingBetweenLines12.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            styleParagraphProperties11.Append(spacingBetweenLines12);

            StyleTableProperties styleTableProperties3 = new StyleTableProperties();
            styleTableProperties3.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            TableIndentation tableIndentation3 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders2 = new TableBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "BFBFBF", ThemeColor = ThemeColorValues.Background1, ThemeShade = "BF", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders2.Append(topBorder2);
            tableBorders2.Append(leftBorder2);
            tableBorders2.Append(bottomBorder2);
            tableBorders2.Append(rightBorder2);
            tableBorders2.Append(insideHorizontalBorder2);
            tableBorders2.Append(insideVerticalBorder2);

            TableCellMarginDefault tableCellMarginDefault3 = new TableCellMarginDefault();
            TopMargin topMargin3 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin3 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin3 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin3 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault3.Append(topMargin3);
            tableCellMarginDefault3.Append(tableCellLeftMargin3);
            tableCellMarginDefault3.Append(bottomMargin3);
            tableCellMarginDefault3.Append(tableCellRightMargin3);

            styleTableProperties3.Append(tableIndentation3);
            styleTableProperties3.Append(tableBorders2);
            styleTableProperties3.Append(tableCellMarginDefault3);

            style24.Append(styleName24);
            style24.Append(basedOn20);
            style24.Append(uIPriority23);
            style24.Append(styleParagraphProperties11);
            style24.Append(styleTableProperties3);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);
            styles1.Append(style9);
            styles1.Append(style10);
            styles1.Append(style11);
            styles1.Append(style12);
            styles1.Append(style13);
            styles1.Append(style14);
            styles1.Append(style15);
            styles1.Append(style16);
            styles1.Append(style17);
            styles1.Append(style18);
            styles1.Append(style19);
            styles1.Append(style20);
            styles1.Append(style21);
            styles1.Append(style22);
            styles1.Append(style23);
            styles1.Append(style24);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15" } };
            fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");

            Font font1 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "A00002EF", UnicodeSignature1 = "4000207B", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of part.
        private void GeneratePartContent(MainDocumentPart part)
        {
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            document1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            document1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            document1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            Body body1 = new Body();

            //Paragraph paragraph1 = new Paragraph();
            //paragraph1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordml");
            //body1.AppendChild(paragraph1);

            document1.Append(body1);

            SectionProperties sectionProperties1 = new SectionProperties();
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U, Orient = PageOrientationValues.Portrait };
            PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1440U, Bottom = 1440, Left = (UInt32Value)1440U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(sectionProperties1);

            part.Document = document1;
            
        }


    }
    internal class CoreProperties
    {
        // Adds child parts and generates content of the specified part.
        public void CreateCoreFilePropertiesPart(CoreFilePropertiesPart part,
            System.Collections.Generic.Dictionary<string,string> dict)
        {
            GeneratePartContent(part,dict);
        }

        // Generates content of part.
        private void GeneratePartContent(CoreFilePropertiesPart part,
            System.Collections.Generic.Dictionary<string, string> dict)
        {
            //System.DateTime currentTime = System.DateTime.UtcNow; // Get the current time in UTC

            //string formattedTime = currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ");
            string title = dict["Title"];// "My Title";//dict["Title"];
            string subject = dict["Subject"];// "My Subject";//dict["Subject"];
            string keywords = dict["Keywords"];//"My Keyword";//dict["Keywords"];
            string description = dict["Description"];// "My Description";//dict["Description"];
            string creator = dict["Creator"];// "FileFormat.Words"; //dict["Creator"];
            string created = dict["Created"];
            string modified = dict["Modified"];
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(part.GetStream(
                System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"utf-8\"?><coreProperties xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"><dc:title>"+ title + "</dc:title><dc:subject>"+ subject +"</dc:subject><keywords>" + keywords + "</keywords><dc:description>"+ description +"</dc:description><dcterms:created xsi:type=\"dcterms:W3CDTF\">" + created +"</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">"+ modified + "</dcterms:modified><dc:creator>"+ creator +"</dc:creator></coreProperties>");
            writer.Flush();
            writer.Close();
        }
    }
    internal class CustomProperties
    {
        // Adds child parts and generates content of the specified part.
        public void CreateExtendedFilePropertiesPart(ExtendedFilePropertiesPart part)
        {
            GeneratePartContent(part);

        }

        // Generates content of part.
        private void GeneratePartContent(ExtendedFilePropertiesPart part)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            properties1.AddNamespaceDeclaration("ap", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "0";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "0";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "0";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "FileFormat.Words";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "0";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "FileFormat.Words";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "0";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "23.10.0";
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal.dotm";

            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);
            properties1.Append(template1);

            part.Properties = properties1;
        }
    }
}
