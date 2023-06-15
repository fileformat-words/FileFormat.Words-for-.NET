using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Reflection.Metadata;


namespace FileFormat.Words
{

    public class Image
    {
        protected internal Drawing drawing;
        protected internal Paragraph parentParagraph;
        //protected internal string caption;

        public Image(Document document, string imagePath, int width, int height)
        {
            MainDocumentPart mainPart = document.wordDocument.MainDocumentPart;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            string imagePartId = mainPart.GetIdOfPart(imagePart);

            this.drawing =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = width * 9525, Cy = height * 9525 },
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = 1U,
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
                                            Id = 0U,
                                            Name = System.IO.Path.GetFileName(imagePath)
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                        { Embed = imagePartId },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = width * 9525, Cy = height * 9525 }),
                                        new A.PresetGeometry(
                                            new A.AdjustValueList()
                                        )
                                        { Preset = A.ShapeTypeValues.Rectangle })))
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                    {
                        DistanceFromTop = 0U,
                        DistanceFromBottom = 0U,
                        DistanceFromLeft = 0U,
                        DistanceFromRight = 0U
                    });

        }

        internal Drawing Drawing
        {
            get
            {
                return this.drawing;
            }
        }

        public long GetExtentCx()
        {
            var inline = this.drawing?.Inline;
            var extent = inline?.Extent;
            return extent?.Cx ?? 0;
        }

        public long GetExtentCy()
        {
            var inline = this.drawing?.Inline;
            var extent = inline?.Extent;
            return extent?.Cy ?? 0;
        }


        public static List<Stream> ExtractImagesFromDocument(Document document)
        {
            MainDocumentPart mainDocumentPart = document.wordDocument.MainDocumentPart;

            List<Stream> imageParts = new List<Stream>();

            foreach (var part in mainDocumentPart.Parts)
            {
                if (part.OpenXmlPart is ImagePart imagePart)
                {
                    imageParts.Add(imagePart.GetStream());
                }
            }

            return imageParts;
        }

        public void AddBorder(string borderColor, int borderWidth)
        {
            // Find the Picture element
            var picture = this.drawing.Inline.Graphic.GraphicData
                .Elements<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().FirstOrDefault();

            if (picture != null)
            {
                // Find the ShapeProperties element
                var shapeProperties = picture.Descendants<PIC.ShapeProperties>().FirstOrDefault();

                if (shapeProperties == null)
                {
                    // Create a new ShapeProperties element if it doesn't exist
                    shapeProperties = new PIC.ShapeProperties();
                    picture.Append(shapeProperties);
                }

                // Create a new Outline element for the border
                var outline = new DocumentFormat.OpenXml.Drawing.Outline(
                    new DocumentFormat.OpenXml.Drawing.SolidFill()
                    {
                        RgbColorModelHex = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = borderColor }
                    },
                    new DocumentFormat.OpenXml.Drawing.Outline { Width = (borderWidth * 12700) } // Convert to EMU (English Metric Unit)
                );

                // Set the border properties in ShapeProperties
                shapeProperties.Append(outline);
            }
        }

        public void RotateImage(int rotationAngle)
        {
            // Find the Picture element
            var picture = this.drawing.Inline.Graphic.GraphicData
                .Elements<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().FirstOrDefault();

            if (picture != null)
            {
                // Find the Transform2D element
                var transform2D = picture.ShapeProperties.Transform2D;

                if (transform2D == null)
                {
                    // Create a new Transform2D element if it doesn't exist
                    transform2D = new A.Transform2D();
                    picture.ShapeProperties.Append(transform2D);
                }

                // Set the rotation angle
                transform2D.Rotation = rotationAngle * 60000;
            }
        }

        public void ResizeImage(int newWidth, int newHeight)
        {
            // Calculate the new extents
            long newCx = newWidth * 9525;
            long newCy = newHeight * 9525;

            // Find the Picture element
            var picture = this.drawing.Inline.Graphic.GraphicData.Elements<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().FirstOrDefault();

            if (picture != null)
            {
                // Update the Extent and Transform2D properties
                picture.ShapeProperties.Transform2D.Extents.Cx = newCx;
                picture.ShapeProperties.Transform2D.Extents.Cy = newCy;
                this.drawing.Inline.Extent.Cx = newCx;
                this.drawing.Inline.Extent.Cy = newCy;
            }
        }

    }
}

