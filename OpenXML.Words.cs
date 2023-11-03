using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DF = DocumentFormat.OpenXml;
using PKG = DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using FF = FileFormat.Words.IElements;
using OWD = OpenXML.Words.Data;
using OT = OpenXML.Templates;

namespace OpenXML.Words
{
    internal class Document
    {
        private PKG.WordprocessingDocument pkgDocument;
        private WP.Body wpBody;
        private MemoryStream ms;
        PKG.MainDocumentPart mainPart;

        internal Document()
        {
            try
            {
                ms = new MemoryStream();
                pkgDocument = PKG.WordprocessingDocument.Create(ms, DF.WordprocessingDocumentType.Document, true);
                mainPart = pkgDocument.AddMainDocumentPart();
                mainPart.Document = new WP.Document();
                OT.DefaultTemplate tmp = new OT.DefaultTemplate();
                tmp.CreateMainDocumentPart(mainPart);
                CreateProperties(pkgDocument);
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Initialize OOXML Element(s)");
                throw new FileFormatException(errorMessage, ex);
            }
        }
        internal void CreateDocument(List<FF.IElement> lst)
        {
            try
            {
                /**pkgDocument = PKG.WordprocessingDocument.Create(ms, DF.WordprocessingDocumentType.Document, true);
                mainPart = pkgDocument.AddMainDocumentPart();
                mainPart.Document = new WP.Document();
                OT.GenerateStructure tmp = new OT.GenerateStructure();
                tmp.CreateMainDocumentPart(mainPart);
                CreateProperties(pkgDocument);**/
                wpBody = mainPart.Document.Body;
                WP.SectionProperties sectionProperties = wpBody.Elements<WP.SectionProperties>().FirstOrDefault();
                int sequence = 1;
                foreach (var element in lst)
                {
                    if (element is FF.Paragraph ffP)
                    {
                        var para = CreateParagraph(ffP);
                        wpBody.InsertBefore(para, sectionProperties);
                        sequence++;
                    }
                    if (element is FF.Image ffImg)
                    {
                        var para = CreateImage(ffImg, mainPart);
                        wpBody.InsertBefore(para, sectionProperties);
                        sequence++;
                    }
                    if (element is FF.Table ffTable)
                    {
                        var table = CreateTable(ffTable);
                        wpBody.InsertBefore(table, sectionProperties);
                        sequence++;
                    }
                }
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Create OOXML Element(s)");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal void CreateProperties(PKG.WordprocessingDocument pkgDocument)
        {
            // Access the core properties part
            var corePart = pkgDocument.CoreFilePropertiesPart;
            if (corePart != null)
            {
                // Delete the core properties part
                pkgDocument.DeletePart(corePart);
            }
            var customPart = pkgDocument.CustomFilePropertiesPart;
            if(customPart != null)
            {
                pkgDocument.DeletePart(customPart);
            }
            OT.CoreProperties coreProperties = new OT.CoreProperties();
            Dictionary<string, string> dictCoreProp = new Dictionary<string, string>();
            dictCoreProp["Title"] = "Newly Created Document";
            dictCoreProp["Subject"] = "WordProcessing Document Generation";
            dictCoreProp["Keywords"] = "DOCX";
            dictCoreProp["Description"] = "A WordProcessing Document Created from Scratch.";
            dictCoreProp["Creator"] = "FileFormat.Words";
            DateTime currentTime = System.DateTime.UtcNow;
            dictCoreProp["Created"] = currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ");
            dictCoreProp["Modified"] = currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ");
            coreProperties.CreateCoreFilePropertiesPart(pkgDocument.AddCoreFilePropertiesPart(), dictCoreProp);
            OT.CustomProperties customProperties = new OT.CustomProperties();
            customProperties.CreateExtendedFilePropertiesPart(pkgDocument.AddExtendedFilePropertiesPart());
        }

        internal WP.Paragraph CreateParagraph(FF.Paragraph ffP)
        {
            try
            {
                var wpParagraph = new WP.Paragraph();

                if (ffP.Style != null)
                {
                    WP.ParagraphProperties paragraphProperties = new WP.ParagraphProperties();
                    WP.ParagraphStyleId paragraphStyleId = new WP.ParagraphStyleId { Val = ffP.Style };
                    paragraphProperties.Append(paragraphStyleId);
                    wpParagraph.Append(paragraphProperties);
                }

                foreach (var ffR in ffP.Runs)
                {
                    var wpRun = new WP.Run();

                    var runProperties = new WP.RunProperties();

                    if (ffR.FontFamily != null)
                    {
                        var runFont = new WP.RunFonts
                        {
                            Ascii = ffR.FontFamily,
                            HighAnsi = ffR.FontFamily,
                            ComplexScript = ffR.FontFamily,
                            EastAsia = ffR.FontFamily
                        };
                        runProperties.Append(runFont);
                    }

                    if (ffR.Color != null)
                    {
                        var color = new WP.Color { Val = ffR.Color };
                        runProperties.Append(color);
                    }

                    if (ffR.FontSize > 0)
                    {
                        var fontSize = new WP.FontSize { Val = (ffR.FontSize *2).ToString() };
                        runProperties.Append(fontSize);
                    }

                    if (ffR.Bold)
                    {
                        runProperties.Append(new WP.Bold() { Val = new DF.OnOffValue(true) });
                    }

                    if (ffR.Italic)
                    {
                        runProperties.Append(new WP.Italic());
                    }

                    if (ffR.Underline)
                    {
                        var underline = new WP.Underline { Val = WP.UnderlineValues.Single };
                        runProperties.Append(underline);
                    }

                    var text = new WP.Text(ffR.Text);
                    wpRun.Append(runProperties, text);
                    wpParagraph.AppendChild(wpRun);
                }
                return wpParagraph;
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Create Paragraph");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal WP.Table CreateTable(FF.Table ffTable)
        {
            try
            {
                int rows = ffTable.Rows.Count;
                int cols = ffTable.Rows[0].Cells.Count;

                WP.Table wpTable = new WP.Table(
                    new WP.TableProperties(
                    new WP.TableStyle() { Val = ffTable.Style } // Specify the TableStyle ID you want to apply
                    )
                );
                WP.TableGrid tableGrid = new WP.TableGrid();
                for (var i = 0; i < cols; i++)
                {
                    if (ffTable.Column.Width > 0)
                        tableGrid.Append(new WP.GridColumn { Width = ffTable.Column.Width.ToString() });
                    else
                        tableGrid.Append(new WP.GridColumn());
                }
                wpTable.Append(tableGrid);

                for (var i = 0; i < rows; i++)
                {
                    WP.TableRow wpRow = new WP.TableRow();

                    for (var j = 0; j < cols; j++)
                    {
                        WP.TableCell wpCell = new WP.TableCell();
                        FF.Cell ffCell = ffTable.Rows[i].Cells[j];
                        foreach (FF.Paragraph ffPara in ffCell.Paragraphs)
                        {
                            wpCell.Append(CreateParagraph(ffPara));
                        }
                        wpRow.Append(wpCell);
                    }
                    wpTable.Append(wpRow);
                }
                return wpTable;
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Create Table");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal WP.Paragraph CreateImage(FF.Image ffIMG,PKG.MainDocumentPart mainPart)
        {
            try
            {
                byte[] imageBytes = ffIMG.ImageData;
                PKG.ImagePart imagePart = mainPart.AddImagePart(PKG.ImagePartType.Png);
                using (Stream partStream = imagePart.GetStream())
                {
                    partStream.Write(imageBytes, 0, imageBytes.Length); // Write the image bytes to the partStream
                }
                float dpi = 96; // The DPI of the image (you may need to adjust this value)
                int widthInPixels;
                int heightInPixels;
                if (ffIMG.Width > 500 || ffIMG.Height > 500)
                {
                    widthInPixels = 500;
                    heightInPixels = 500;
                }
                else if(ffIMG.Width == 0 || ffIMG.Height == 300)
                {
                    widthInPixels = 500;
                    heightInPixels = 500;
                }
                else
                {
                    widthInPixels = ffIMG.Width;
                    heightInPixels = ffIMG.Height;
                }
                float widthInInches = widthInPixels / dpi ;
                float heightInInches = heightInPixels / dpi;

                long widthInEMU = (long)(widthInInches * 914400);
                long heightInEMU = (long)(heightInInches * 914400);
                //long widthInEMU = (long)widthInInches;
                //long heightInEMU = (long)heightInInches;

                // Define the reference of the image.
                var element =
                     new WP.Drawing(
                         new DW.Inline(
                             //new DW.Extent() { Cx = ffIMG.Width*9525 , Cy = ffIMG.Height*9525 },
                             new DW.Extent() { Cx = widthInEMU, Cy = heightInEMU },
                             new DW.EffectExtent()
                             {
                                 LeftEdge = 0L,
                                 TopEdge = 0L,
                                 RightEdge = 0L,
                                 BottomEdge = 0L
                             },
                             new DW.DocProperties()
                             {
                                 Id = (DF.UInt32Value)1U,
                                 Name = "Picture 1"
                             },
                             new DW.NonVisualGraphicFrameDrawingProperties(
                                 new A.GraphicFrameLocks() { NoChangeAspect = true }),
                             new A.Graphic(
                                 new A.GraphicData(
                                     new PIC.Picture(
                                         new PIC.NonVisualPictureProperties(
                                             new PIC.NonVisualDrawingProperties()
                                             {
                                                 Id = (DF.UInt32Value)0U,
                                                 Name = "New Bitmap Image.jpg"
                                             },
                                             new PIC.NonVisualPictureDrawingProperties()),
                                         new PIC.BlipFill(
                                             new A.Blip(
                                                 new A.BlipExtensionList(
                                                     new A.BlipExtension()
                                                     {
                                                         Uri =
                                                            "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                     })
                                             )
                                             {
                                                 Embed = mainPart.GetIdOfPart(imagePart),
                                                 CompressionState =
                                                 A.BlipCompressionValues.Print
                                             },
                                             new A.Stretch(
                                                 new A.FillRectangle())),
                                         new PIC.ShapeProperties(
                                             new A.Transform2D(
                                                 new A.Offset() { X = 0L, Y = 0L },
                                                 //new A.Extents() { Cx = ffIMG.Width*9525, Cy = ffIMG.Height*9525 }
                                                 new A.Extents() { Cx = widthInEMU, Cy = heightInEMU }),
                                             new A.PresetGeometry(
                                                 new A.AdjustValueList()
                                             )
                                             { Preset = A.ShapeTypeValues.Rectangle }))
                                 )
                                 { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                         )
                         {
                             DistanceFromTop = (DF.UInt32Value)0U,
                             DistanceFromBottom = (DF.UInt32Value)0U,
                             DistanceFromLeft = (DF.UInt32Value)0U,
                             DistanceFromRight = (DF.UInt32Value)0U,
                             EditId = "50D07946"
                         });
                return new WP.Paragraph(new WP.Run(element));
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Create Image");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal List<FF.IElement> LoadDocument(Stream stream)
        {
            try
            {
                /**
                using (var fs = new FileStream(filename, FileMode.Open))
                {
                    fs.CopyTo(ms);
                }**/
                //stream.CopyTo(ms);

                pkgDocument = PKG.WordprocessingDocument.Open(stream, true);

                OWD.OOXMLDocData.SetPKG(pkgDocument);

                mainPart = pkgDocument.MainDocumentPart;
                wpBody = pkgDocument.MainDocumentPart.Document.Body;

                int sequence = 1;
                var lstelement = new List<FF.IElement>();

                foreach (var element in wpBody.Elements())
                {
                    if (element is WP.Paragraph wpPara)
                    {
                        bool drawingFound = false;
                        foreach (var drawing in wpPara.Descendants<WP.Drawing>())
                        {
                            FF.Image image = LoadImage(drawing, sequence);
                            if (image != null)
                            {
                                lstelement.Add(LoadImage(drawing, sequence));
                                sequence++;
                                drawingFound = true;
                            }
                            else
                            {
                                lstelement.Add(new FF.Unknown { ElementID = sequence });
                                sequence++;
                                drawingFound = true;
                            }
                        }
                        if (!drawingFound)
                        {
                            lstelement.Add(LoadParagraph(wpPara, sequence));
                            sequence++;
                        }
                    }
                    else if (element is WP.Drawing drawing)
                    {
                        FF.Image image = LoadImage(drawing, sequence);
                        if (image != null)
                        {
                            lstelement.Add(LoadImage(drawing, sequence));
                            sequence++;
                        }
                        else
                        {
                            lstelement.Add(new FF.Unknown { ElementID = sequence });
                            sequence++;
                        }
                    }
                    else if (element is WP.Table wpTable)
                    {
                        lstelement.Add(LoadTable(wpTable, sequence));
                        sequence++;
                    }
                    else if (element is WP.SectionProperties wpSection)
                    {
                        lstelement.Add(LoadSection(wpSection, sequence));
                        sequence++;
                    }
                    else
                    {
                        lstelement.Add(new FF.Unknown { ElementID = sequence });
                        sequence++;
                    }
                }
                return lstelement;
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Load OOXML Elements");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal FF.Paragraph LoadParagraph(WP.Paragraph wpPara, int ID)
        {
            try
            {
                var ffP = new FF.Paragraph { ElementID = ID };

                WP.ParagraphProperties paraProps = wpPara.GetFirstChild<WP.ParagraphProperties>();
                if (paraProps != null)
                {
                    WP.ParagraphStyleId paraStyleId = paraProps.Elements<WP.ParagraphStyleId>().FirstOrDefault();
                    if (paraStyleId != null)
                    {
                        ffP.Style = paraStyleId.Val.Value;
                    }
                }

                var runs = wpPara.Elements<WP.Run>();
                var lstR = new List<FF.Run>();

                foreach (var wpR in runs)
                {
                    int? fontSize = wpR.RunProperties?.FontSize?.Val != null ? int.Parse(wpR.RunProperties.FontSize.Val) : (int?)null;
                    if (fontSize != null) fontSize /= 2;
                    var ffR = new FF.Run
                    {
                        Text = wpR.InnerText,
                        FontFamily = wpR.RunProperties?.RunFonts?.Ascii ?? null,
                        FontSize = fontSize ?? 0,//int.Parse(wpR.RunProperties?.FontSize?.Val ?? null),
                        Color = wpR.RunProperties?.Color?.Val ?? null,
                        Bold = (wpR.RunProperties != null && wpR.RunProperties.Bold != null),
                        Italic = (wpR.RunProperties != null && wpR.RunProperties.Italic != null),
                        Underline = (wpR.RunProperties != null && wpR.RunProperties.Underline != null)
                    };
                    ffP.AddRun(ffR);
                }
                return ffP;
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Load Paragraph");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal FF.Image LoadImage(WP.Drawing drawing, int sequence)
        {
            try
            {
                FF.Image image;
                foreach (A.Blip blip in drawing.Descendants<A.Blip>())
                {
                    if (blip != null)
                    {
                        DW.Extent extent = drawing.Inline.Extent;

                        if (extent != null)
                        {

                            int dpi = 96; // Replace with your image's DPI

                            int widthInPixels = (int)(extent.Cx / (914400 / dpi));
                            int heightInPixels = (int)(extent.Cy / (914400 / dpi));

                            PKG.ImagePart imagePart = mainPart.GetPartById(blip.Embed) as PKG.ImagePart;
                            if (imagePart != null)
                            {
                                using (Stream stream = imagePart.GetStream())
                                {
                                    image = new FF.Image();
                                    image.ElementID = sequence;
                                    byte[] imageBytes;
                                    using (MemoryStream memoryStream = new MemoryStream())
                                    {
                                        stream.CopyTo(memoryStream);
                                        imageBytes = memoryStream.ToArray();
                                    }
                                    image.ImageData = imageBytes;

                                    image.Height = heightInPixels;
                                    image.Width = widthInPixels;
                                    return image;
                                }
                            }
                        }
                    }

                }
                return null;
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Load Image");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal FF.Table LoadTable(WP.Table wpTable,int ID)
        {
            try
            {
                List<FF.Row> lstffRows = new List<FF.Row>();
                foreach (WP.TableRow wpRow in wpTable.Elements<WP.TableRow>())
                {
                    FF.Row ffRow = new FF.Row();
                    ffRow.Cells = new List<FF.Cell>();
                    foreach (WP.TableCell wpCell in wpRow.Elements<WP.TableCell>())
                    {
                        List<FF.Paragraph> ffParas = new List<FF.Paragraph>();
                        foreach (WP.Paragraph paragraph in wpCell.Elements<WP.Paragraph>())
                        {
                            ffParas.Add(LoadParagraph(paragraph, 0));
                        }
                        FF.Cell ffCell = new FF.Cell { Paragraphs = ffParas };
                        ffRow.Cells.Add(ffCell);
                    }
                    lstffRows.Add(ffRow);
                }
                FF.Table ffTable = new FF.Table();
                ffTable.Rows = lstffRows;
                ffTable.ElementID = ID;
                WP.TableGrid tableGrid = wpTable.Elements<WP.TableGrid>().FirstOrDefault();
                if (tableGrid != null)
                {
                    WP.GridColumn gridColumn = tableGrid.Elements<WP.GridColumn>().FirstOrDefault();
                    ffTable.Column.Width = Convert.ToInt32(gridColumn.Width);
                }
                var tableProperties = wpTable.Descendants<WP.TableProperties>().FirstOrDefault();
                if (tableProperties != null)
                {
                    var tableStyle = tableProperties.TableStyle;
                    if (tableStyle != null)
                    {
                        ffTable.Style = tableStyle.Val;
                    }
                }
                return ffTable;
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Load Table");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal FF.Section LoadSection(WP.SectionProperties sectPr,int ID)
        {
            try
            {
                FF.Section section = new FF.Section();
                section.ElementID = ID;
                if (sectPr != null)
                {
                    WP.PageSize pageSize = sectPr.Elements<WP.PageSize>().FirstOrDefault();
                    if (pageSize != null)
                    {
                        section.PageSize = new FF.PageSize
                        {
                            Height = int.Parse(pageSize.Height),
                            Width = int.Parse(pageSize.Width),
                            Orientation = pageSize.Orient,
                        };
                    }

                    WP.PageMargin pageMargin = sectPr.Elements<WP.PageMargin>().FirstOrDefault();
                    if (pageMargin != null)
                    {
                        section.PageMargin = new FF.PageMargin
                        {
                            Top = int.Parse(pageMargin.Top),
                            Right = int.Parse(pageMargin.Right),
                            Bottom = int.Parse(pageMargin.Bottom),
                            Left = int.Parse(pageMargin.Left),
                            Header = int.Parse(pageMargin.Header),
                            Footer = int.Parse(pageMargin.Footer),
                        };
                    }
                }

                return section;
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Load Section");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal FF.ElementStyles LoadStyles()
        {
            try
            {
                FF.ElementStyles elementStyles = new FF.ElementStyles();
                var themePart = mainPart.ThemePart;
                if (themePart != null)
                {
                    var theme = themePart.Theme;
                    foreach (var fontScheme in theme.Elements())
                    {
                        foreach (var latinFont in fontScheme.Descendants<A.LatinFont>())
                        {
                            elementStyles.ThemeFonts.Add(latinFont.Typeface);
                        }
                    }
                    foreach (var fontScheme in theme.Elements())
                    {
                        var fonts = fontScheme.Descendants<A.SupplementalFont>();

                        foreach (var font in fonts)
                        {
                            if (font.Typeface != null)
                            {
                                elementStyles.ThemeFonts.Add(font.Typeface);
                            }
                        }
                    }
                }
                var fontTablePart = mainPart.FontTablePart;
                if (fontTablePart != null)
                {
                    var fontTable = fontTablePart.Fonts.Elements<WP.Font>();

                    foreach (var font in fontTable)
                    {
                        elementStyles.TableFonts.Add(font.Name);
                    }
                }
                var styleDefinitionsPart = mainPart.StyleDefinitionsPart;

                if (styleDefinitionsPart != null)
                {
                    WP.Styles styles = styleDefinitionsPart.Styles;
                    if (styles != null)
                    {
                        foreach (WP.Style style in styles.Elements<WP.Style>())
                        {
                            if (style.Type != null && style.Type == WP.StyleValues.Paragraph)
                            {
                                elementStyles.ParagraphStyles.Add(style.StyleId);
                            }
                            if (style.Type != null && style.Type == WP.StyleValues.Table)
                            {
                                elementStyles.TableStyles.Add(style.StyleId);
                            }
                        }
                    }
                }
                return elementStyles;
            }
            catch(Exception ex)
            {
                return null;
            }
        }

        internal void SaveDocument(Stream stream)
        {
            try
            {
                pkgDocument.Clone(stream);
                pkgDocument.Dispose();
                ms.Dispose();
            }
            catch(Exception ex)
            {
                string errorMessage = OWD.OOXMLDocData.ConstructMessage(ex, "Save OOXML Document");
                throw new FileFormatException(errorMessage, ex);
            }
        }
    }
}

