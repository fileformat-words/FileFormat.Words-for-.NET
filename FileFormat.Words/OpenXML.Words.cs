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
    internal class OwDocument
    {
        private PKG.WordprocessingDocument _pkgDocument;
        private WP.Body _wpBody;
        private MemoryStream _ms;
        private PKG.MainDocumentPart _mainPart;
        private readonly object _lockObject = new object();
        private OwDocument()
        {
            lock (_lockObject)
            {
                try
                {
                    _ms = new MemoryStream();
                    _pkgDocument = PKG.WordprocessingDocument.Create(_ms, DF.WordprocessingDocumentType.Document, true);
                    _mainPart = _pkgDocument.AddMainDocumentPart();
                    _mainPart.Document = new WP.Document();
                    var tmp = new OT.DefaultTemplate();
                    tmp.CreateMainDocumentPart(_mainPart);
                    AddNumberingDefinitions(_pkgDocument);
                    CreateProperties(_pkgDocument);
                }
                catch (Exception ex)
                {
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Initialize OOXML Element(s)");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        public static OwDocument CreateInstance()
        {
            return new OwDocument(); 
        }

        internal void CreateDocument(List<FF.IElement> lst)
        {
            try
            {
                _wpBody = _mainPart.Document.Body;
                if (_wpBody == null)
                    throw new FileFormatException("Package or Document or Body is null", new NullReferenceException());
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
                        case FF.Table ffTable:
                        {
                            var table = CreateTable(ffTable);
                            _wpBody.InsertBefore(table, sectionProperties);
                            break;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Create OOXML Element(s)");
                throw new FileFormatException(errorMessage, ex);
            }
        }

        internal void CreateProperties(PKG.WordprocessingDocument pkgDocument)
        {
            var corePart = pkgDocument.CoreFilePropertiesPart;
            if (corePart != null)
            {
                pkgDocument.DeletePart(corePart);
            }
            var customPart = pkgDocument.CustomFilePropertiesPart;
            if(customPart != null)
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
                        
                        WP.JustificationValues justificationValue = MapAlignmentToJustification(ffP.Alignment);
                        paragraphProperties.Append(new WP.Justification { Val = justificationValue });
                        

                        if (ffP.Indentation != null)
                        {
                            SetIndentation(paragraphProperties, ffP.Indentation);
                        }

                        if (ffP.IsNumbered || ffP.IsBullet)
                        {
                            var numberingProperties = new WP.NumberingProperties();
                            var numberingId = new WP.NumberingId();
                            var numberingLevelReference = new WP.NumberingLevelReference();

                            if (ffP.IsBullet)
                            {                                
                                numberingId.Val = 1;
                                numberingLevelReference.Val = ffP.NumberingLevel ?? 0;
                            }
                            else if (ffP.IsNumbered)
                            {
                                numberingId.Val = ffP.NumberingId <= 1 || ffP.NumberingId == null ? 2 : ffP.NumberingId;
                                numberingLevelReference.Val = ffP.NumberingLevel ?? 0;
                            }

                            numberingProperties.Append(numberingId, numberingLevelReference);
                            paragraphProperties.Append(numberingProperties);
                        }
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
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal void AddNumberingDefinitions(PKG.WordprocessingDocument pkgDocument)
        {
            PKG.NumberingDefinitionsPart numberingPart = pkgDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = pkgDocument.MainDocumentPart.AddNewPart<PKG.NumberingDefinitionsPart>();
            }

            WP.Numbering numbering = new WP.Numbering();

            WP.AbstractNum abstractNumBulleted = new WP.AbstractNum() { AbstractNumberId = 1 };
            WP.AbstractNum abstractNumNumbered = new WP.AbstractNum() { AbstractNumberId = 2 };

            for (int i = 0; i < 9; i++)
            {
                abstractNumBulleted.Append(CreateLevel(i, WP.NumberFormatValues.Bullet, "•"));
                abstractNumNumbered.Append(CreateLevel(i, WP.NumberFormatValues.Decimal, $"%{i + 1}."));
            }

            numbering.Append(abstractNumBulleted);
            numbering.Append(abstractNumNumbered);

            WP.NumberingInstance numInstanceBulleted = new WP.NumberingInstance() { NumberID = 1 };
            numInstanceBulleted.Append(new WP.AbstractNumId() { Val = abstractNumBulleted.AbstractNumberId });

            WP.NumberingInstance numInstanceNumbered = new WP.NumberingInstance() { NumberID = 2 };
            numInstanceNumbered.Append(new WP.AbstractNumId() { Val = abstractNumNumbered.AbstractNumberId });

            numbering.Append(numInstanceBulleted);
            numbering.Append(numInstanceNumbered);

            numberingPart.Numbering = numbering;
        }

        private WP.Level CreateLevel(int levelIndex, WP.NumberFormatValues numFormatVal, string levelTextVal)
        {
            WP.Level level = new WP.Level(
                new WP.StartNumberingValue() { Val = 1 },
                new WP.NumberingFormat() { Val = numFormatVal },
                new WP.LevelText() { Val = levelTextVal },
                new WP.LevelJustification() { Val = WP.LevelJustificationValues.Left }
            )
            { LevelIndex = levelIndex };
            if (numFormatVal == WP.NumberFormatValues.Bullet)
            {
                level.RemoveAllChildren<WP.StartNumberingValue>();
            }
            return level;
        }

        private void SetIndentation(WP.ParagraphProperties paragraphProperties, FF.Indentation ffIndentation)
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

        private WP.JustificationValues MapAlignmentToJustification(FF.ParagraphAlignment alignment)
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
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal WP.Paragraph CreateImage(FF.Image ffImg,PKG.MainDocumentPart mainPart)
        {
            lock (_lockObject)
            {
                try
                {
                    var imageBytes = ffImg.ImageData;
                    var imagePart = mainPart.AddImagePart(PKG.ImagePartType.Png);
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
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal List<FF.IElement> LoadDocument(Stream stream)
        {
            lock (_lockObject)
            {
                try
                {
                    _pkgDocument = PKG.WordprocessingDocument.Open(stream, true);

                    if (_pkgDocument.MainDocumentPart?.Document?.Body == null) throw new FileFormatException("Package or Document or Body is null", new NullReferenceException());

                    OWD.OoxmlDocData.CreateInstance(_pkgDocument);

                    _mainPart = _pkgDocument.MainDocumentPart;
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
                                        elements.Add(LoadImage(drawing, sequence));
                                        sequence++;
                                        drawingFound = true;
                                    }
                                    else
                                    {
                                        elements.Add(new FF.Unknown { ElementId = sequence });
                                        sequence++;
                                        drawingFound = true;
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
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

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
                    var numberingProperties = wpPara.ParagraphProperties?.NumberingProperties;
                    if (numberingProperties != null)
                    {
                        var numIdVal = numberingProperties.NumberingId?.Val;
                        var levelVal = numberingProperties.NumberingLevelReference?.Val;

                        // Check if there's a valid numbering id and level reference
                        if (numIdVal.HasValue && levelVal.HasValue)
                        {
                            // Get the numbering part from the document
                            var numberingPart = _pkgDocument.MainDocumentPart.NumberingDefinitionsPart;
                            if (numberingPart != null)
                            {
                                // Look for the AbstractNum that matches the numIdVal
                                var abstractNum = numberingPart.Numbering.Elements<WP.AbstractNum>()
                                                .FirstOrDefault(a => a.AbstractNumberId.Value == numIdVal.Value);

                                if (abstractNum != null)
                                {
                                    // Get the level corresponding to the levelVal
                                    var level = abstractNum.Elements<WP.Level>()
                                                    .FirstOrDefault(l => l.LevelIndex.Value == levelVal.Value);

                                    // If the level's numbering format is bullet, set IsBullet to true
                                    if (level != null && level.NumberingFormat != null && level.NumberingFormat.Val.Value == WP.NumberFormatValues.Bullet)
                                    {
                                        ffP.IsBullet = true;
                                        ffP.NumberingLevel = levelVal.Value;
                                    }
                                    // If the level's numbering format is not bullet, it's a numbered list
                                    else if (level != null)
                                    {
                                        ffP.IsNumbered = true;
                                        ffP.NumberingLevel = levelVal.Value;
                                        ffP.NumberingId = numIdVal.Value;
                                    }
                                }
                            }
                        }
                    }
                    var justificationElement = paraProps.Elements<WP.Justification>().FirstOrDefault();
                    if (justificationElement != null)
                    {
                        ffP.Alignment = MapJustificationToAlignment(justificationElement.Val);
                    }
                    var Indentation = paraProps.Elements<WP.Indentation>().FirstOrDefault();
                    if (Indentation != null) { 
                        if (Indentation.Left != null)
                        {
                            ffP.Indentation.Left = int.Parse(Indentation.Left);                        
                        }
                        if (Indentation.Right != null)
                        {
                            ffP.Indentation.Right = int.Parse(Indentation.Right);
                        }
                        if (Indentation.Hanging != null)
                        {
                            ffP.Indentation.Hanging = int.Parse(Indentation.Hanging);
                        }
                        if (Indentation.FirstLine != null)
                        {                        
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
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        private bool IsBulletStyle(string styleId)
        {            
            return styleId == "BulletStyle";
        }

        private FF.ParagraphAlignment MapJustificationToAlignment(WP.JustificationValues justificationValue)
        {
            switch (justificationValue)
            {
                case WP.JustificationValues.Left:
                    return FF.ParagraphAlignment.Left;
                case WP.JustificationValues.Center:
                    return FF.ParagraphAlignment.Center;
                case WP.JustificationValues.Right:
                    return FF.ParagraphAlignment.Right;
                case WP.JustificationValues.Both:
                    return FF.ParagraphAlignment.Justify;
                default:
                    return FF.ParagraphAlignment.Left;
            }
        }

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

                                var imagePart = _mainPart.GetPartById(blip.Embed) as PKG.ImagePart;
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
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal FF.Table LoadTable(WP.Table wpTable,int id)
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
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal FF.Section LoadSection(WP.SectionProperties sectPr,int id)
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
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

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
                    var errorMessage = OWD.OoxmlDocData.ConstructMessage(ex, "Save OOXML OWDocument");
                    throw new FileFormatException(errorMessage, ex);
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
    }
}

