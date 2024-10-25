using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DF = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DWG = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using DWS = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using FF = FileFormat.Words.IElements;
using OWD = OpenXML.Words.Data;
using OT = OpenXML.Templates;
using FileFormat.Words;

namespace OpenXML.Words
{
    internal class OwDocument
    {
        private WordprocessingDocument _pkgDocument;
        private WP.Body _wpBody;
        private MemoryStream _ms;
        private MainDocumentPart _mainPart;
        private List<int> _IDs;
        private NumberingDefinitionsPart _numberingPart;
        private readonly object _lockObject = new object();
        private OwDocument()
        {
            lock (_lockObject)
            {
                try
                {
                    _ms = new MemoryStream();
                    _pkgDocument = WordprocessingDocument.Create(_ms, DF.WordprocessingDocumentType.Document, true);
                    _mainPart = _pkgDocument.AddMainDocumentPart();
                    _mainPart.Document = new WP.Document();
                    var tmp = new OT.DefaultTemplate();
                    tmp.CreateMainDocumentPart(_mainPart);
                    CreateProperties(_pkgDocument);

                    _numberingPart = _mainPart.NumberingDefinitionsPart;

                    if (_numberingPart != null)
                    {
                        _IDs = new List<int>();
                        foreach (var abstractNum in _numberingPart.Numbering.Elements<WP.AbstractNum>())
                        {
                            _IDs.Add(abstractNum.AbstractNumberId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Initialize OOXML Element(s)");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }
        private OwDocument(WordprocessingDocument pkg)
        {
            lock (_lockObject)
            {
                try
                {
                    //_ms = new MemoryStream();
                    _pkgDocument = pkg;
                    _mainPart = pkg.MainDocumentPart;
                    //_mainPart.Document = new WP.Document();
                    //var tmp = new OT.DefaultTemplate();
                    //tmp.CreateMainDocumentPart(_mainPart);
                    //CreateProperties(_pkgDocument);

                    _numberingPart = _mainPart.NumberingDefinitionsPart;

                    if (_numberingPart != null)
                    {
                        _IDs = new List<int>();
                        foreach (var abstractNum in _numberingPart.Numbering.Elements<WP.AbstractNum>())
                        {
                            _IDs.Add(abstractNum.AbstractNumberId);
                        }
                    }
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Initialize OOXML Element(s)");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }

        #region Create Core Properties for OpenXML Word Document
        internal void CreateProperties(WordprocessingDocument pkgDocument)
        {
            var corePart = pkgDocument.CoreFilePropertiesPart;
            if (corePart != null)
            {
                pkgDocument.DeletePart(corePart);
            }
            var customPart = pkgDocument.CustomFilePropertiesPart;
            if (customPart != null)
            {
                pkgDocument.DeletePart(customPart);
            }
            var coreProperties = new OT.CoreProperties();
            var dictCoreProp = new Dictionary<string, string>
            {
                ["Title"] = "Newly Created OWDocument",
                ["Subject"] = "WordProcessing OWDocument Generation",
                ["Keywords"] = "DOCX",
                ["Description"] = "A WordProcessing OWDocument Created from Scratch.",
                ["Creator"] = "FileFormat.Words"
            };
            var currentTime = System.DateTime.UtcNow;
            dictCoreProp["Created"] = currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ");
            dictCoreProp["Modified"] = currentTime.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ");
            coreProperties.CreateCoreFilePropertiesPart(pkgDocument.AddCoreFilePropertiesPart(), dictCoreProp);
            var customProperties = new OT.CustomProperties();
            customProperties.CreateExtendedFilePropertiesPart(pkgDocument.AddExtendedFilePropertiesPart());
        }
        #endregion

        public static OwDocument CreateInstance()
        {
            return new OwDocument();
        }

        public static OwDocument CreateInstance(WordprocessingDocument pkg)
        {
            return new OwDocument(pkg);
        }

        #region Create OpenXML Word Document Contents Based on FileFormat.Words.IElements

        #region Main Method
        internal void CreateDocument(List<FF.IElement> lst)
        {
            try
            {
                _wpBody = _mainPart.Document.Body;

                if (_wpBody == null)
                    throw new FileFormat.Words.FileFormatException("Package or Document or Body is null", new NullReferenceException());

                var sectionProperties = _wpBody.Elements<WP.SectionProperties>().FirstOrDefault();

                foreach (var element in lst)
                {
                    switch (element)
                    {
                        case FF.Paragraph ffP:
                            {
                                var para = CreateParagraph(ffP);
                                _wpBody.InsertBefore(para, sectionProperties);
                                break;
                            }
                        case FF.Image ffImg:
                            {
                                var para = CreateImage(ffImg, _mainPart);
                                _wpBody.InsertBefore(para, sectionProperties);
                                break;
                            }

                        case FF.Shape ffShape:
                            {
                                var para = CreateShape(ffShape);
                                _wpBody.InsertBefore(para, sectionProperties);
                                break;
                            }

                        case FF.GroupShape ffGroupShape:
                            {
                                var para = CreateConnectedShapes(ffGroupShape);
                                _wpBody.InsertBefore(para, sectionProperties);
                                break;
                            }

                        case FF.Table ffTable:
                            {
                                var table = CreateTable(ffTable);
                                _wpBody.InsertBefore(table, sectionProperties);
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Initialize OOXML Element(s)");
                throw new FileFormat.Words.FileFormatException(errorMessage, ex);
            }

        }
        #endregion

        #region Create OpenXML Paragraph
        internal WP.Paragraph CreateParagraph(FF.Paragraph ffP)
        {
            lock (_lockObject)
            {
                try
                {
                    var wpParagraph = new WP.Paragraph();

                    if (ffP.Style != null)
                    {
                        var paragraphProperties = new WP.ParagraphProperties();
                        
                        var paragraphStyleId = new WP.ParagraphStyleId { Val = ffP.Style };
                        paragraphProperties.Append(paragraphStyleId);
                        
                        #region Create List Paragraph

                        if (ffP.Style == "ListParagraph")
                        {

                            // Check if NumberingId already exists

                            var isExist = false;
                            if (_IDs != null)
                            {
                                foreach (var id in _IDs)
                                {
                                    if (id == ffP.NumberingId)
                                    {
                                        isExist = true;

                                        var numbering = _numberingPart.Numbering;
                                        var abstractNum = numbering.Elements<WP.AbstractNum>().FirstOrDefault(an => an.AbstractNumberId == ffP.NumberingId);
                                        if (abstractNum != null)
                                        {
                                            var level = abstractNum.Elements<WP.Level>().FirstOrDefault(l => l.LevelIndex == ffP.NumberingLevel - 1);
                                            if (level != null)
                                            {
                                                if (ffP.IsAlphabeticNumber)
                                                {
                                                    level.NumberingFormat.Val = WP.NumberFormatValues.LowerLetter;
                                                    level.LevelText.Val = string.Format("%{0}.", (int)ffP.NumberingLevel);
                                                }
                                                else if (ffP.IsRoman)
                                                {
                                                    level.NumberingFormat.Val = WP.NumberFormatValues.LowerRoman;
                                                    level.LevelText.Val = string.Format("%{0}.", (int)ffP.NumberingLevel);
                                                }
                                                else if (ffP.IsBullet)
                                                {
                                                    level.NumberingFormat.Val = WP.NumberFormatValues.Bullet;
                                                    level.LevelText.Val = "o";
                                                }
                                                if (ffP.IsNumbered)
                                                {
                                                    level.NumberingFormat.Val = WP.NumberFormatValues.Decimal;
                                                    level.LevelText.Val = string.Format("%{0}.", string.Join(".%", Enumerable.Range(1, (int)ffP.NumberingLevel)));
                                                }
                                                numbering.Save();
                                            }
                                        }
                                    }
                                }
                            }
                            if (!isExist)
                            {
                                if (ffP.NumberingId != null)
                                {
                                    if (ffP.NumberingLevel == null)
                                        ffP.NumberingLevel = 1;
                                    if (ffP.IsAlphabeticNumber == false && ffP.IsBullet == false &&
                                        ffP.IsNumbered == false && ffP.IsRoman == false)
                                        ffP.IsNumbered = true;

                                    var abstractNum = new WP.AbstractNum() { AbstractNumberId = ffP.NumberingId };
                                    abstractNum.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                    var multiLevelType = new WP.MultiLevelType() { Val = WP.MultiLevelValues.Multilevel };
                                    multiLevelType.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                    abstractNum.Append(multiLevelType);

                                    var level = new WP.Level() { LevelIndex = 0 };

                                    for (var i = 1; i <= 9; i++)
                                    {
                                        level = new WP.Level() { LevelIndex = i - 1 };
                                        level.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                                        var numberingFormat = new WP.NumberingFormat();
                                        var levelText = new WP.LevelText();

                                        if (ffP.IsNumbered)
                                        {
                                            numberingFormat.Val = WP.NumberFormatValues.Decimal;
                                            levelText.Val = string.Format("%{0}.", string.Join(".%", Enumerable.Range(1, i)));
                                        }
                                        else if (ffP.IsAlphabeticNumber)
                                        {
                                            numberingFormat.Val = WP.NumberFormatValues.LowerLetter;
                                            levelText.Val = string.Format("%{0}.", i);
                                        }
                                        else if (ffP.IsRoman)
                                        {
                                            numberingFormat.Val = WP.NumberFormatValues.LowerRoman;
                                            levelText.Val = string.Format("%{0}.", i);
                                        }
                                        else if (ffP.IsBullet)
                                        {
                                            numberingFormat.Val = WP.NumberFormatValues.Bullet;
                                            levelText.Val = "o";
                                        }

                                        var previousParagraphProperties = new WP.PreviousParagraphProperties();
                                        var indentation = new WP.Indentation() { Left = (i * 720).ToString(), Hanging = "360" };
                                        previousParagraphProperties.Append(indentation);

                                        level.Append(new WP.StartNumberingValue() { Val = 1 });
                                        level.Append(numberingFormat);
                                        level.Append(levelText);
                                        level.Append(new WP.LevelJustification() { Val = WP.LevelJustificationValues.Left });
                                        level.Append(previousParagraphProperties);

                                        abstractNum.Append(level);
                                    }

                                    var numberingInstance = new WP.NumberingInstance() { NumberID = ffP.NumberingId };
                                    var abstractNumId = new WP.AbstractNumId() { Val = ffP.NumberingId };

                                    numberingInstance.Append(abstractNumId);
                                    _numberingPart.Numbering.Append(abstractNum);
                                    _numberingPart.Numbering.Append(numberingInstance);


                                    _IDs.Add((int)ffP.NumberingId);
                                }
                            }

                            var numberingProperties = new WP.NumberingProperties();
                            var numberingLevelReference = new WP.NumberingLevelReference() { Val = ffP.NumberingLevel - 1 };
                            var numberingId = new WP.NumberingId() { Val = ffP.NumberingId };
                            numberingProperties.Append(numberingLevelReference);
                            numberingProperties.Append(numberingId);
                            paragraphProperties.Append(numberingProperties);
                        }
                        #endregion


                        // Create Borders
                        if (ffP.ParagraphBorder.Size > 0)
                        {
                            WP.ParagraphBorders paragraphBorders = new WP.ParagraphBorders();
                            WP.TopBorder topBorder = new WP.TopBorder()
                            { Val = CreateBorder(ffP.ParagraphBorder.Width),
                                Color = ffP.ParagraphBorder.Color,
                                Size = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size,
                                Space = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size
                            };
                            WP.LeftBorder leftBorder = new WP.LeftBorder()
                            {
                                Val = CreateBorder(ffP.ParagraphBorder.Width),
                                Color = ffP.ParagraphBorder.Color,
                                Size = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size,
                                Space = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size
                            };
                            WP.BottomBorder bottomBorder = new WP.BottomBorder()
                            {
                                Val = CreateBorder(ffP.ParagraphBorder.Width),
                                Color = ffP.ParagraphBorder.Color,
                                Size = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size,
                                Space = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size
                            };
                            WP.RightBorder rightBorder = new WP.RightBorder()
                            {
                                Val = CreateBorder(ffP.ParagraphBorder.Width),
                                Color = ffP.ParagraphBorder.Color,
                                Size = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size,
                                Space = (DF.UInt32Value)(uint)ffP.ParagraphBorder.Size
                            };

                            paragraphBorders.Append(topBorder);
                            paragraphBorders.Append(leftBorder);
                            paragraphBorders.Append(bottomBorder);
                            paragraphBorders.Append(rightBorder);

                            paragraphProperties.Append(paragraphBorders);
                        }
                        // Create Justification
                        WP.JustificationValues justificationValue = CreateJustification(ffP.Alignment);
                        paragraphProperties.Append(new WP.Justification { Val = justificationValue });

                        // Create Indentation
                        CreateIndentation(paragraphProperties, ffP.Indentation);

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
                            var fontSize = new WP.FontSize { Val = (ffR.FontSize * 2).ToString() };
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
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Paragraph");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }

        private WP.JustificationValues CreateJustification(FF.ParagraphAlignment alignment)
        {
            switch (alignment)
            {
                case FF.ParagraphAlignment.Left:
                    return WP.JustificationValues.Left;
                case FF.ParagraphAlignment.Center:
                    return WP.JustificationValues.Center;
                case FF.ParagraphAlignment.Right:
                    return WP.JustificationValues.Right;
                case FF.ParagraphAlignment.Justify:
                    return WP.JustificationValues.Both;
                default:
                    return WP.JustificationValues.Left;
            }
        }

        private WP.BorderValues CreateBorder(FF.BorderWidth borderWidth)
        {
            switch (borderWidth)
            {
                case FF.BorderWidth.Single:
                    return WP.BorderValues.Single;
                case FF.BorderWidth.Double:
                    return WP.BorderValues.Double;
                case FF.BorderWidth.Dotted:
                    return WP.BorderValues.Dotted;
                case FF.BorderWidth.DotDash:
                    return WP.BorderValues.DotDash;
                default:
                    return WP.BorderValues.Single;
            }
        }

        private void CreateIndentation(WP.ParagraphProperties paragraphProperties, FF.Indentation ffIndentation)
        {
            var indentation = new WP.Indentation();

            if (ffIndentation.Left > 0)
            {
                indentation.Left = (ffIndentation.Left * 1440).ToString();
            }

            if (ffIndentation.Right > 0)
            {
                indentation.Right = (ffIndentation.Right * 1440).ToString();
            }

            if (ffIndentation.FirstLine > 0)
            {
                indentation.FirstLine = (ffIndentation.FirstLine * 1440).ToString();
            }

            if (ffIndentation.Hanging > 0)
            {
                indentation.Hanging = (ffIndentation.Hanging * 1440).ToString();
            }

            paragraphProperties.Append(indentation);
        }
        #endregion

        #region Create OpenXML Table
        internal WP.Table CreateTable(FF.Table ffTable)
        {
            lock (_lockObject)
            {
                try
                {
                    var rows = ffTable.Rows.Count;
                    var cols = ffTable.Rows[0].Cells.Count;

                    var wpTable = new WP.Table(
                        new WP.TableProperties(
                            new WP.TableStyle() { Val = ffTable.Style } // Specify the TableStyle ID you want to apply
                        )
                    );
                    var tableGrid = new WP.TableGrid();
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
                        var wpRow = new WP.TableRow();

                        for (var j = 0; j < cols; j++)
                        {
                            var wpCell = new WP.TableCell();
                            var ffCell = ffTable.Rows[i].Cells[j];
                            foreach (var ffPara in ffCell.Paragraphs)
                            {
                                wpCell.Append(CreateParagraph(ffPara));
                            }

                            wpRow.Append(wpCell);
                        }

                        wpTable.Append(wpRow);
                    }

                    return wpTable;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Table");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }
        #endregion

        #region Create OpenXML Image
        internal WP.Paragraph CreateImage(FF.Image ffImg, MainDocumentPart mainPart)
        {
            lock (_lockObject)
            {
                try
                {
                    var imageBytes = ffImg.ImageData;
                    var imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    using (var partStream = imagePart.GetStream())
                    {
                        partStream.Write(imageBytes, 0, imageBytes.Length); // Write the image bytes to the partStream
                    }

                    float dpi = 96; // The DPI of the image (you may need to adjust this value)
                    //int widthInPixels;
                    //int heightInPixels;
                    const int maxDimension = 500;

                    var widthInPixels = (ffImg.Width > 0 && ffImg.Width <= maxDimension) ? ffImg.Width : maxDimension;
                    var heightInPixels = (ffImg.Height > 0 && ffImg.Height <= maxDimension) ? ffImg.Height : maxDimension;

                    var widthInInches = widthInPixels / dpi;
                    var heightInInches = heightInPixels / dpi;

                    var widthInEmu = (long)(widthInInches * 914400);
                    var heightInEmu = (long)(heightInInches * 914400);
                    //long widthInEMU = (long)widthInInches;
                    //long heightInEMU = (long)heightInInches;

                    // Define the reference of the image.
                    var element =
                        new WP.Drawing(
                            new DW.Inline(
                                //new DW.Extent() { Cx = ffIMG.Width*9525 , Cy = ffIMG.Height*9525 },
                                new DW.Extent() { Cx = widthInEmu, Cy = heightInEmu },
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
                                                        new A.Extents() { Cx = widthInEmu, Cy = heightInEmu }),
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
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Image");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }
        #endregion

        #region Create OpenXML Shape
        internal WP.Paragraph CreateShape(FF.Shape shape)
        {
            lock (_lockObject)
            {
                try
                {
                    var paragraph = new WP.Paragraph();
                    var run = new WP.Run();

                    var runProperties = new WP.RunProperties();
                    var noProof = new WP.NoProof();

                    runProperties.Append(noProof);

                    var alternateContent = new DF.AlternateContent();
                    alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    var alternateContentChoice = new DF.AlternateContentChoice() { Requires = "wps" };
                    alternateContentChoice.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    var drawing = new WP.Drawing();
                    drawing.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                    var inline = new DW.Inline()
                    { DistanceFromTop = (DF.UInt32Value)0U, DistanceFromBottom = (DF.UInt32Value)0U, DistanceFromLeft = (DF.UInt32Value)0U, DistanceFromRight = (DF.UInt32Value)0U, AnchorId = "27EE2959", EditId = "551435BE" };
                    inline.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                    inline.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");

                    var extent = new DW.Extent() { Cx = shape.Width * 9525, Cy = shape.Height * 9525 };
                    extent.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var effectExtent = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 13970L, BottomEdge = 13970L };
                    effectExtent.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var docProperties = new DW.DocProperties() { Id = (DF.UInt32Value)1609145151U, Name = "Oval 1" };
                    docProperties.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var nonVisualGraphicFrameDrawingProperties = new DW.NonVisualGraphicFrameDrawingProperties();
                    nonVisualGraphicFrameDrawingProperties.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var graphic = new A.Graphic();
                    graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                    var graphicData = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

                    var wordprocessingShape = new DWS.WordprocessingShape();
                    wordprocessingShape.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                    var nonVisualDrawingShapeProperties = new DWS.NonVisualDrawingShapeProperties();

                    var shapeProperties = new DWS.ShapeProperties();

                    var transform2D = new A.Transform2D();
                    var offset = new A.Offset() { X = shape.X * 9525, Y = shape.Y * 9525 };
                    var extents = new A.Extents() { Cx = shape.Width * 9525, Cy = shape.Height * 9525 };

                    transform2D.Append(offset);
                    transform2D.Append(extents);

                    var presetGeometry = new A.PresetGeometry() { Preset = CreateShapeType(shape.Type) }; //A.ShapeTypeValues.Ellipse };
                    var adjustValueList = new A.AdjustValueList();

                    presetGeometry.Append(adjustValueList);
                    var outline = new A.Outline();

                    shapeProperties.Append(transform2D);
                    shapeProperties.Append(presetGeometry);
                    shapeProperties.Append(outline);

                    var shapeStyle = new DWS.ShapeStyle();

                    var lineReference = new A.LineReference() { Index = (DF.UInt32Value)2U };

                    var schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
                    var shade = new A.Shade() { Val = 50000 };

                    schemeColor.Append(shade);

                    lineReference.Append(schemeColor);

                    var fillReference = new A.FillReference() { Index = (DF.UInt32Value)1U };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

                    fillReference.Append(schemeColor);

                    var effectReference = new A.EffectReference() { Index = (DF.UInt32Value)0U };
                    var rgbColorModelPercentage = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

                    effectReference.Append(rgbColorModelPercentage);

                    var fontReference = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

                    fontReference.Append(schemeColor);

                    shapeStyle.Append(lineReference);
                    shapeStyle.Append(fillReference);
                    shapeStyle.Append(effectReference);
                    shapeStyle.Append(fontReference);
                    var textBodyProperties = new DWS.TextBodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };

                    wordprocessingShape.Append(nonVisualDrawingShapeProperties);
                    wordprocessingShape.Append(shapeProperties);
                    wordprocessingShape.Append(shapeStyle);
                    wordprocessingShape.Append(textBodyProperties);

                    graphicData.Append(wordprocessingShape);

                    graphic.Append(graphicData);

                    inline.Append(extent);
                    inline.Append(effectExtent);
                    inline.Append(docProperties);
                    inline.Append(nonVisualGraphicFrameDrawingProperties);
                    inline.Append(graphic);

                    drawing.Append(inline);

                    alternateContentChoice.Append(drawing);

                    var alternateContentFallback = new DF.AlternateContentFallback();
                    alternateContentFallback.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    alternateContent.Append(alternateContentChoice);
                    alternateContent.Append(alternateContentFallback);

                    run.Append(runProperties);
                    run.Append(alternateContent);

                    paragraph.Append(run);

                    return paragraph;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Shape");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }

        private A.ShapeTypeValues CreateShapeType(FF.ShapeType shapeType)
        {
            switch (shapeType)
            {
                case FF.ShapeType.Rectangle:
                    return A.ShapeTypeValues.Rectangle;
                case FF.ShapeType.Triangle:
                    return A.ShapeTypeValues.Triangle;
                case FF.ShapeType.Ellipse:
                    return A.ShapeTypeValues.Ellipse;
                case FF.ShapeType.Diamond:
                    return A.ShapeTypeValues.Diamond;
                case FF.ShapeType.Hexagone:
                    return A.ShapeTypeValues.Hexagon;
                default:
                    return A.ShapeTypeValues.Ellipse;
            }
        }

        #region "Shapes with connector"
        internal WP.Paragraph CreateConnectedShapes(FF.GroupShape groupShape)
        {
            lock (_lockObject)
            {
                try
                {
                    if (groupShape.Shape2.X < (groupShape.Shape1.X+ groupShape.Shape1.Width))
                        throw new FileFormat.Words.FileFormatException("Invalid shape dimensions",
                            new ArgumentException());

                    var paragraph = new WP.Paragraph();
                    paragraph.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordml");

                    var run = new WP.Run();

                    var runProperties = new WP.RunProperties();
                    var noProof = new WP.NoProof();

                    runProperties.Append(noProof);

                    var alternateContent = new DF.AlternateContent();
                    alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    var alternateContentChoice = new DF.AlternateContentChoice() { Requires = "wpg" };
                    alternateContentChoice.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    var drawing = new WP.Drawing();
                    drawing.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

                    var inline = new DW.Inline() { DistanceFromTop = (DF.UInt32Value)0U, DistanceFromBottom = (DF.UInt32Value)0U, DistanceFromLeft = (DF.UInt32Value)0U, DistanceFromRight = (DF.UInt32Value)0U, AnchorId = "24C249F3", EditId = "163BC827" };
                    inline.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                    inline.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");

                    var extent = new DW.Extent() { Cx = 3778250L, Cy = 622300L };
                    extent.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var effectExtent = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 12700L, BottomEdge = 25400L };
                    effectExtent.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var docProperties = new DW.DocProperties() { Id = (DF.UInt32Value)122768519U, Name = "Group-" + groupShape.ElementId.ToString() };
                    docProperties.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var nonVisualGraphicFrameDrawingProperties = new DW.NonVisualGraphicFrameDrawingProperties();
                    nonVisualGraphicFrameDrawingProperties.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

                    var graphic = new A.Graphic();
                    graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

                    A.GraphicData graphicData = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" };

                    var wordprocessingGroup = new DWG.WordprocessingGroup();
                    wordprocessingGroup.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                    var nonVisualGroupDrawingShapeProperties = new DWG.NonVisualGroupDrawingShapeProperties();

                    var shape1 = groupShape.Shape1;
                    var shape2 = groupShape.Shape2;

                    var groupShapeProperties = new DWG.GroupShapeProperties();

                    var transformGroup = new A.TransformGroup();

                    /****** Group Offsets ******/
                    // var offset = new A.Offset() { X = 0L, Y = 0L };
                    // Group.X=shape1.x, Group.Y=shape1.y
                    var groupX = shape1.X * 9525;
                    var groupY = shape1.Y * 9525;
                    var groupOffset = new A.Offset()
                    {
                        X = groupX,
                        Y = groupY
                    };

                    /****** Group Extents ******/
                    // var extents = new A.Extents() { Cx = 3778250L, Cy = 622300L };
                    // Group.Width=278(shape2.X)-0(shape1.X)+118(shape2.Width)
                    var groupCx = (shape2.X - shape1.X + shape2.Width) * 9525;
                    // Group.Height=Shape1.Height
                    var groupCy = shape1.Height * 9525;
                    var groupExtents = new A.Extents() { Cx = groupCx, Cy = groupCy };

                    /****** Child Offset & Extents ******/
                    //var childOffset = new A.ChildOffset() { X=0L, Y=0L};
                    // Same as group.X and group.Y
                    var childOffset = new A.ChildOffset()
                    {
                        X = groupX,
                        Y = groupY
                    };
                    //var childExtents = new A.ChildExtents() { Cx = 3778250L, Cy = 622300L };
                    var childExtents = new A.ChildExtents()
                    {
                        Cx = groupCx,
                        Cy = groupCy
                    };

                    transformGroup.Append(groupOffset);
                    transformGroup.Append(groupExtents);
                    transformGroup.Append(childOffset);
                    transformGroup.Append(childExtents);

                    groupShapeProperties.Append(transformGroup);

                    /******************* shapes ****************/
                    var wordprocessingShape01 = CreatePartialShape(
                        shape1.ElementId, shape1.X, shape1.Y,
                        shape1.Width, shape1.Height, CreateShapeType(shape1.Type)
                        );
                    var wordprocessingShape02 = CreatePartialShape(
                        shape2.ElementId, shape2.X, shape2.Y,
                        shape2.Width, shape2.Height, CreateShapeType(shape2.Type)
                        );

                    /**************** connector *****************/
                    // Connector
                    var wordprocessingShape = new DWS.WordprocessingShape();
                    wordprocessingShape.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
                    var nonVisualDrawingProperties = new DWS.NonVisualDrawingProperties()
                    {
                        Id = (DF.UInt32Value)161453463U,
                        Name = "Connector: Elbow 161453463"
                    };

                    var nonVisualConnectorProperties = new DWS.NonVisualConnectorProperties();
                    //A.StartConnection startConnection1 = new A.StartConnection() { Id = (DF.UInt32Value)448142074U, Index = (DF.UInt32Value)3U };
                    A.StartConnection startConnection = new A.StartConnection()
                    {
                        Id = (DF.UInt32Value)(uint)shape1.ElementId,
                        Index = (DF.UInt32Value)3U
                    };
                    //A.EndConnection endConnection1 = new A.EndConnection() { Id = (DF.UInt32Value)1011268246U, Index = (DF.UInt32Value)2U };
                    A.EndConnection endConnection = new A.EndConnection()
                    {
                        Id = (DF.UInt32Value)(uint)shape2.ElementId,
                        Index = (DF.UInt32Value)2U
                    };

                    nonVisualConnectorProperties.Append(startConnection);
                    nonVisualConnectorProperties.Append(endConnection);

                    var shapeProperties = new DWS.ShapeProperties();

                    var transform2D = new A.Transform2D();
                    //var offset4 = new A.Offset() { X = 914400L, Y = 311150L };
                    // 96 (same as shape1.width),33 (half of shape1.height)
                    var connectorX = shape1.Width * 9525;
                    var connectorY = shape1.Height / 2 * 9525;
                    var offset4 = new A.Offset()
                    {
                        X = connectorX,
                        Y = connectorY
                    };
                    //var extents4 = new A.Extents() { Cx = 1733550L, Cy = 6350L };
                    // 182, 1
                    // connector.Cx = Group.Width - (shape1.Width+shape2.Width)
                    var connectorCx = groupCx - (connectorX + (shape2.Width * 9525));
                    var extents4 = new A.Extents() { Cx = connectorCx, Cy = 6350L };

                    transform2D.Append(offset4);
                    transform2D.Append(extents4);

                    var presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.BentConnector3 };
                    var adjustValueList = new A.AdjustValueList();

                    presetGeometry.Append(adjustValueList);
                    var outline = new A.Outline();

                    shapeProperties.Append(transform2D);
                    shapeProperties.Append(presetGeometry);
                    shapeProperties.Append(outline);

                    var shapeStyle = new DWS.ShapeStyle();

                    var lineReference = new A.LineReference() { Index = (DF.UInt32Value)1U };
                    var schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

                    lineReference.Append(schemeColor);

                    var fillReference = new A.FillReference() { Index = (DF.UInt32Value)0U };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

                    fillReference.Append(schemeColor);

                    var effectReference = new A.EffectReference() { Index = (DF.UInt32Value)0U };
                    A.RgbColorModelPercentage rgbColorModelPercentage = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

                    effectReference.Append(rgbColorModelPercentage);

                    var fontReference = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Text1 };

                    fontReference.Append(schemeColor);

                    shapeStyle.Append(lineReference);
                    shapeStyle.Append(fillReference);
                    shapeStyle.Append(effectReference);
                    shapeStyle.Append(fontReference);
                    var textBodyProperties = new DWS.TextBodyProperties();

                    wordprocessingShape.Append(nonVisualDrawingProperties);
                    wordprocessingShape.Append(nonVisualConnectorProperties);
                    wordprocessingShape.Append(shapeProperties);
                    wordprocessingShape.Append(shapeStyle);
                    wordprocessingShape.Append(textBodyProperties);

                    wordprocessingGroup.Append(nonVisualGroupDrawingShapeProperties);
                    wordprocessingGroup.Append(groupShapeProperties);
                    wordprocessingGroup.Append(wordprocessingShape01);
                    wordprocessingGroup.Append(wordprocessingShape02);
                    wordprocessingGroup.Append(wordprocessingShape);

                    graphicData.Append(wordprocessingGroup);

                    graphic.Append(graphicData);

                    inline.Append(extent);
                    inline.Append(effectExtent);
                    inline.Append(docProperties);
                    inline.Append(nonVisualGraphicFrameDrawingProperties);
                    inline.Append(graphic);

                    drawing.Append(inline);

                    alternateContentChoice.Append(drawing);

                    var alternateContentFallback = new DF.AlternateContentFallback();
                    alternateContentFallback.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");

                    alternateContent.Append(alternateContentChoice);
                    alternateContent.Append(alternateContentFallback);

                    run.Append(runProperties);
                    run.Append(alternateContent);

                    paragraph.Append(run);

                    return paragraph;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Coonected Shapes");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }

        private DWS.WordprocessingShape CreatePartialShape(int Id, int X, int Y, int Width, int Height, A.ShapeTypeValues shapeTypeValues)
        {
            lock (_lockObject)
            {
                try
                {
                    var wordprocessingShape = new DWS.WordprocessingShape();
                    wordprocessingShape.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

                    //var nonVisualDrawingProperties1 = new DWS.NonVisualDrawingProperties() { Id = (DF.UInt32Value)448142074U, Name = "Rectangle 448142074" };
                    var nonVisualDrawingProperties = new DWS.NonVisualDrawingProperties()
                    {
                        Id = (DF.UInt32Value)(uint)Id,
                        Name = "shape-" + Id.ToString()
                    };
                    var nonVisualDrawingShapeProperties = new DWS.NonVisualDrawingShapeProperties();

                    var shapeProperties = new DWS.ShapeProperties();
                    var transform2D = new A.Transform2D();
                    //var offset = new A.Offset() { X = 0L, Y = 0L };
                    var offset = new A.Offset() { X = X * 9525, Y = Y * 9525 };
                    //var extents = new A.Extents() { Cx = 914400L, Cy = 622300L };
                    var extents = new A.Extents() { Cx = Width * 9525, Cy = Height * 9525 };

                    transform2D.Append(offset);
                    transform2D.Append(extents);

                    //var presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
                    var presetGeometry = new A.PresetGeometry()
                    {
                        Preset = shapeTypeValues
                    };
                    var adjustValueList = new A.AdjustValueList();

                    presetGeometry.Append(adjustValueList);
                    var outline = new A.Outline();

                    shapeProperties.Append(transform2D);
                    shapeProperties.Append(presetGeometry);
                    shapeProperties.Append(outline);

                    var shapeStyle = new DWS.ShapeStyle();

                    var lineReference = new A.LineReference()
                    {
                        Index = (DF.UInt32Value)2U
                    };

                    var schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };
                    var shade = new A.Shade() { Val = 50000 };

                    schemeColor.Append(shade);

                    lineReference.Append(schemeColor);

                    var fillReference = new A.FillReference() { Index = (DF.UInt32Value)1U };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Accent1 };

                    fillReference.Append(schemeColor);

                    var effectReference = new A.EffectReference() { Index = (DF.UInt32Value)0U };
                    var rgbColorModelPercentage = new A.RgbColorModelPercentage() { RedPortion = 0, GreenPortion = 0, BluePortion = 0 };

                    effectReference.Append(rgbColorModelPercentage);

                    var fontReference = new A.FontReference() { Index = A.FontCollectionIndexValues.Minor };
                    schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.Light1 };

                    fontReference.Append(schemeColor);

                    shapeStyle.Append(lineReference);
                    shapeStyle.Append(fillReference);
                    shapeStyle.Append(effectReference);
                    shapeStyle.Append(fontReference);

                    var textBodyProperties = new DWS.TextBodyProperties() { Anchor = A.TextAnchoringTypeValues.Center };

                    wordprocessingShape.Append(nonVisualDrawingProperties);
                    wordprocessingShape.Append(nonVisualDrawingShapeProperties);
                    wordprocessingShape.Append(shapeProperties);
                    wordprocessingShape.Append(shapeStyle);
                    wordprocessingShape.Append(textBodyProperties);

                    return wordprocessingShape;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create Partial Shape");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }

        #endregion

        #endregion

        #endregion

        #region Load OpenXML Word Document Content into FileFormat.Words.IElements

        #region Main Method
        internal List<FF.IElement> LoadDocument(Stream stream)
        {
            lock (_lockObject)
            {
                try
                {
                    _pkgDocument = WordprocessingDocument.Open(stream, true);

                    if (_pkgDocument.MainDocumentPart?.Document?.Body == null) throw new FileFormat.Words.FileFormatException("Package or Document or Body is null", new NullReferenceException());

                    OWD.OoxmlDocData.CreateInstance(_pkgDocument);

                    _mainPart = _pkgDocument.MainDocumentPart;
                    _numberingPart = _mainPart.NumberingDefinitionsPart;
                    _wpBody = _pkgDocument.MainDocumentPart.Document.Body;

                    var sequence = 1;
                    var elements = new List<FF.IElement>();

                    foreach (var element in _wpBody.Elements())
                    {

                        switch (element)
                        {
                            case WP.Paragraph wpPara:
                                {
                                    var drawingFound = false;

                                    foreach (var drawing in wpPara.Descendants<WP.Drawing>())
                                    {                                   
                                        var image = LoadImage(drawing, sequence);

                                        if (image != null)
                                        {
                                            elements.Add(image);
                                            sequence++;
                                            drawingFound = true;
                                        }
                                        else
                                        {
                                            var shape = LoadShape(drawing,sequence);
                                            if (shape != null)
                                            {
                                                elements.Add(shape);
                                                sequence++;
                                                drawingFound = true;
                                            }
                                        }
                                    }

                                    if (!drawingFound)
                                    {
                                        elements.Add(LoadParagraph(wpPara, sequence));
                                        sequence++;
                                    }

                                    break;
                                }

                            case WP.Drawing drawing:
                                {

                                    var image = LoadImage(drawing, sequence);
                                    if (image != null)
                                    {
                                        elements.Add(LoadImage(drawing, sequence));
                                        sequence++;
                                    }
                                    else
                                    {
                                        elements.Add(new FF.Unknown { ElementId = sequence });
                                        sequence++;
                                    }

                                    break;
                                }
                            case WP.Table wpTable:
                                elements.Add(LoadTable(wpTable, sequence));
                                sequence++;
                                break;

                            case WP.SectionProperties wpSection:
                                elements.Add(LoadSection(wpSection, sequence));
                                sequence++;
                                break;
                            default:
                                elements.Add(new FF.Unknown { ElementId = sequence });
                                sequence++;
                                break;
                        }

                    }

                    return elements;
                
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load OOXML Elements");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }

        #endregion

        #region Load OpenXML Paragraph
        internal FF.Paragraph LoadParagraph(WP.Paragraph wpPara, int id)
        {
            lock (_lockObject)
            {
                try
                {
                    var ffP = new FF.Paragraph { ElementId = id };

                    var paraProps = wpPara.GetFirstChild<WP.ParagraphProperties>();
                    if (paraProps != null)
                    {
                        var paraStyleId = paraProps.Elements<WP.ParagraphStyleId>().FirstOrDefault();
                        if (paraStyleId != null)
                        {
                            if (paraStyleId.Val != null) ffP.Style = paraStyleId.Val.Value;
                        }
                    }

                    if (ffP.Style == "ListParagraph")
                    {
                        if (isNumbered(paraProps))
                        {
                            if (_numberingPart != null)
                            {
                                if (paraProps.NumberingProperties.NumberingId.Val != null &&
                                paraProps.NumberingProperties.NumberingLevelReference.Val != null)
                                {
                                    ffP.NumberingId = paraProps.NumberingProperties.NumberingId.Val;
                                    ffP.NumberingLevel = paraProps.NumberingProperties.NumberingLevelReference.Val + 1;

                                    var numbering = _numberingPart.Numbering;
                                    var abstractNum = numbering.Elements<WP.AbstractNum>().FirstOrDefault(an => an.AbstractNumberId == ffP.NumberingId);
                                    if (abstractNum != null)
                                    {
                                        var level = abstractNum.Elements<WP.Level>().FirstOrDefault(l => l.LevelIndex == ffP.NumberingLevel - 1);
                                        if (level != null)
                                        {
                                            if (level.NumberingFormat.Val == WP.NumberFormatValues.Decimal)
                                                ffP.IsNumbered = true;
                                            else if (level.NumberingFormat.Val == WP.NumberFormatValues.LowerLetter)
                                                ffP.IsAlphabeticNumber = true;
                                            else if (level.NumberingFormat.Val == WP.NumberFormatValues.LowerRoman)
                                                ffP.IsRoman = true;
                                            else if (level.NumberingFormat.Val == WP.NumberFormatValues.Bullet)
                                                ffP.IsBullet = true;
                                            else
                                                ffP.IsNumbered = true;
                                        }
                                    }
                                }

                            }
                        }
                    }

                    // Load Border
                    if(isBordered(paraProps))
                    {
                        var topBorder = paraProps.ParagraphBorders?.TopBorder;
                        if (topBorder != null)
                            {
                            ffP.ParagraphBorder.Width = LoadBorder(topBorder.Val);
                            ffP.ParagraphBorder.Color = topBorder.Color;
                            ffP.ParagraphBorder.Size = (int)(uint)topBorder.Size;
                        }
                    }

                    // Load Justification
                    if (isJustified(paraProps))
                    {
                        var justificationElement = paraProps.Elements<WP.Justification>().FirstOrDefault();
                        if (justificationElement != null)
                            ffP.Alignment = LoadAlignment(justificationElement.Val);
                    }
                    else ffP.Alignment = FF.ParagraphAlignment.Left;

                    // Load Indentation
                    if (isIndented(paraProps))
                    {
                        var Indentation = paraProps.Elements<WP.Indentation>().FirstOrDefault();
                        if (Indentation != null)
                        {
                            if (Indentation.Left != null)
                                ffP.Indentation.Left = int.Parse(Indentation.Left);
                            if (Indentation.Right != null)
                                ffP.Indentation.Right = int.Parse(Indentation.Right);
                            if (Indentation.Hanging != null)
                                ffP.Indentation.Hanging = int.Parse(Indentation.Hanging);
                            if (Indentation.FirstLine != null)
                                ffP.Indentation.FirstLine = int.Parse(Indentation.FirstLine);
                        }
                    }

                    var runs = wpPara.Elements<WP.Run>();

                    foreach (var wpR in runs)
                    {
                        var fontSize = wpR.RunProperties?.FontSize?.Val != null
                            ? int.Parse(wpR.RunProperties.FontSize.Val)
                            : (int?)null;
                        if (fontSize != null) fontSize /= 2;
                        var ffR = new FF.Run
                        {
                            Text = wpR.InnerText,
                            FontFamily = wpR.RunProperties?.RunFonts?.Ascii ?? null,
                            FontSize = fontSize ?? 0,
                            Color = wpR.RunProperties?.Color?.Val ?? null,
                            Bold = (wpR.RunProperties != null && wpR.RunProperties.Bold != null),
                            Italic = (wpR.RunProperties != null && wpR.RunProperties.Italic != null),
                            Underline = (wpR.RunProperties != null && wpR.RunProperties.Underline != null)
                        };
                        ffP.AddRun(ffR);
                    }

                    return ffP;
                }

                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Paragraph");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }

        private FF.ParagraphAlignment LoadAlignment(WP.JustificationValues justificationValue)
        {
            if (justificationValue == WP.JustificationValues.Left)
                return FF.ParagraphAlignment.Left;
            else if (justificationValue == WP.JustificationValues.Center)
                return FF.ParagraphAlignment.Center;
            else if (justificationValue == WP.JustificationValues.Right)
                return FF.ParagraphAlignment.Right;
            else if (justificationValue == WP.JustificationValues.Both)
                return FF.ParagraphAlignment.Justify;
            else
                return FF.ParagraphAlignment.Left;
        }

        private FF.BorderWidth LoadBorder(WP.BorderValues borderValue)
        {
            if (borderValue == WP.BorderValues.Single)
                return FF.BorderWidth.Single;
            else if (borderValue == WP.BorderValues.Double)
                return FF.BorderWidth.Double;
            else if (borderValue == WP.BorderValues.Dotted)
                return FF.BorderWidth.Dotted;
            else if (borderValue == WP.BorderValues.DotDash)
                return FF.BorderWidth.DotDash;
            else
                return FF.BorderWidth.Single;
        }

        private bool isBordered(WP.ParagraphProperties prop)
        {
            try
            {
                var paragraphBorders = prop.ParagraphBorders;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool isJustified(WP.ParagraphProperties prop)
        {
            try
            {
                var justification = prop.Justification;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool isIndented(WP.ParagraphProperties prop)
        {
            try
            {
                var indentation = prop.Indentation;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool isNumbered(WP.ParagraphProperties prop)
        {
            try
            {
                var numbering = prop.NumberingProperties;
                var numberingId = numbering.NumberingId.Val;
                var numberingRef = numbering.NumberingLevelReference.Val;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        #endregion

        #region Load OpenXML Image
        internal FF.Image LoadImage(WP.Drawing drawing, int sequence)
        {
            lock (_lockObject)
            {
                try
                {
                    foreach (var blip in drawing.Descendants<A.Blip>())
                    {
                        if (blip != null)
                        {
                            var extent = drawing.Inline.Extent;

                            if (extent != null)
                            {
                                var dpi = 96; // Replace with your image's DPI

                                var widthInPixels = (int)(extent.Cx / (914400 / dpi));
                                var heightInPixels = (int)(extent.Cy / (914400 / dpi));

                                var imagePart = _mainPart.GetPartById(blip.Embed) as ImagePart;
                                if (imagePart == null) continue;
                                using var stream = imagePart.GetStream();
                                var image = new FF.Image
                                {
                                    ElementId = sequence
                                };
                                byte[] imageBytes;
                                using (var memoryStream = new MemoryStream())
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

                    return null;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Image");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }
        #endregion

        internal FF.Shape LoadShape(WP.Drawing drawing, int sequence)
        {
            var inline = drawing.Inline;

            if (inline != null)
            {
                // Extract shape information from inline
                var graphic = inline.Graphic;
                var graphicData = graphic.GraphicData;

                if (graphicData.Uri.Value == "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
                {
                    var wordprocessingShape = graphicData.GetFirstChild<DWS.WordprocessingShape>();
                    if (wordprocessingShape != null)
                    {
                        // Extract position and size from shape properties
                        var shapeProperties = wordprocessingShape.GetFirstChild<DWS.ShapeProperties>();
                        var transform2D = shapeProperties.GetFirstChild<A.Transform2D>();

                        var offset = transform2D.Offset;
                        var extents = transform2D.Extents;

                        int x = (int)(offset.X.Value / 9525); // Convert EMU to points
                        int y = (int)(offset.Y.Value / 9525);
                        int width = (int)(extents.Cx.Value / 9525);
                        int height = (int)(extents.Cy.Value / 9525);

                        // Determine the shape type
                        var presetGeometry = shapeProperties.GetFirstChild<A.PresetGeometry>();
                        var shapeType = FF.ShapeType.Ellipse; // Default

                        if (presetGeometry.Preset == A.ShapeTypeValues.Diamond)
                        {
                            shapeType = FF.ShapeType.Diamond;
                        }
                        else if (presetGeometry.Preset == A.ShapeTypeValues.Hexagon)
                        {
                            shapeType = FF.ShapeType.Hexagone;
                        }
                        var shape = new FF.Shape(x, y, width, height, shapeType);
                        shape.ElementId = sequence;

                        // Return the shape object with extracted data
                        return shape;
                    }
                }
            }
            return null;
        }

        #region Load OpenXML Table
        internal FF.Table LoadTable(WP.Table wpTable, int id)
        {
            lock (_lockObject)
            {
                try
                {
                    var ffRows = new List<FF.Row>();
                    foreach (var wpRow in wpTable.Elements<WP.TableRow>())
                    {
                        var ffRow = new FF.Row
                        {
                            Cells = new List<FF.Cell>()
                        };
                        foreach (var wpCell in wpRow.Elements<WP.TableCell>())
                        {
                            var ffParas = new List<FF.Paragraph>();
                            foreach (var paragraph in wpCell.Elements<WP.Paragraph>())
                            {
                                ffParas.Add(LoadParagraph(paragraph, 0));
                            }

                            var ffCell = new FF.Cell { Paragraphs = ffParas };
                            ffRow.Cells.Add(ffCell);
                        }

                        ffRows.Add(ffRow);
                    }

                    var ffTable = new FF.Table
                    {
                        Rows = ffRows,
                        ElementId = id
                    };
                    var tableGrid = wpTable.Elements<WP.TableGrid>().FirstOrDefault();
                    if (tableGrid != null)
                    {
                        var gridColumn = tableGrid.Elements<WP.GridColumn>().FirstOrDefault();
                        ffTable.Column.Width = Convert.ToInt32(gridColumn.Width);
                    }

                    var tableProperties = wpTable.Descendants<WP.TableProperties>().FirstOrDefault();
                    if (tableProperties == null) return ffTable;
                    var tableStyle = tableProperties.TableStyle;
                    if (tableStyle != null)
                    {
                        ffTable.Style = tableStyle.Val;
                    }

                    return ffTable;
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Table");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }
        #endregion

        #region Load OpenXML Section
        internal FF.Section LoadSection(WP.SectionProperties sectPr, int id)
        {
            lock (_lockObject)
            {
                try
                {
                    var section = new FF.Section
                    {
                        ElementId = id
                    };
                    if (sectPr != null)
                    {
                        var pageSize = sectPr.Elements<WP.PageSize>().FirstOrDefault();
                        if (pageSize != null)
                        {
                            section.PageSize = new FF.PageSize
                            {
                                Height = int.Parse(pageSize.Height),
                                Width = int.Parse(pageSize.Width),
                                Orientation = pageSize.Orient,
                            };
                        }

                        var pageMargin = sectPr.Elements<WP.PageMargin>().FirstOrDefault();
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
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Load Section");
                    throw new FileFormat.Words.FileFormatException(errorMessage, ex);
                }
            }
        }
        #endregion

        #region Load OpenXML Styles
        internal FF.ElementStyles LoadStyles()
        {
            lock (_lockObject)
            {
                try
                {
                    var elementStyles = new FF.ElementStyles();
                    var themePart = _mainPart.ThemePart;
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

                    var fontTablePart = _mainPart.FontTablePart;
                    if (fontTablePart != null)
                    {
                        var fontTable = fontTablePart.Fonts.Elements<WP.Font>();

                        foreach (var font in fontTable)
                        {
                            elementStyles.TableFonts.Add(font.Name);
                        }
                    }

                    var styleDefinitionsPart = _mainPart.StyleDefinitionsPart;

                    if (styleDefinitionsPart == null) return elementStyles;
                    var styles = styleDefinitionsPart.Styles;
                    if (styles != null)
                    {
                        foreach (var style in styles.Elements<WP.Style>())
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

                    return elementStyles;
                }
                catch (Exception ex)
                {
                    return null;
                }
            }
        }
        #endregion

        #endregion

        #region Save OpenXML Word Document to Stream
        internal void SaveDocument(Stream stream)
        {
            lock (_lockObject)
            {
                try
                {
                    _pkgDocument.Clone(stream);
                    //_pkgDocument.Dispose();
                    //_ms.Dispose();
                }
                catch (Exception ex)
                {
                    //var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Save OOXML OWDocument");
                    //throw new FileFormatException(errorMessage, ex);
                    throw new Exception(ex.Message);
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose of managed resources (if any)
                if (_pkgDocument != null)
                {
                    _pkgDocument.Dispose();
                    _pkgDocument = null;
                }
            }
            // Dispose of unmanaged resources
            if (_ms == null) return;
            _ms.Dispose();
            _ms = null;
        }
        #endregion
    }
}
