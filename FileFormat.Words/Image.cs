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
    /// <summary>
    /// This class contains methods to add, view and edit images into a Word document.
    /// </summary>
    public class Image
    {
        /// <value>
        /// An object of the Parent Drawing class.
        /// </value>
        protected internal Drawing drawing;

        /// <value>
        /// An object of the Parent Paragraph class.
        /// </value>
        protected internal Paragraph parentParagraph;
        //protected internal string caption;

        /// <summary>
        /// Initialize an object of the Image class that inserts an image into a Word document.
        /// </summary>
        /// <param name="document">An object of the Document.</param>
        /// <param name="imagePath">String value that represents the path of an image file.</param>
        /// <param name="width">An integer value that represents the image's width.</param>
        /// <param name="height">An integer value that represents the image's height.</param>
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
        /// <summary>
        /// It returns an object of the Drawing class.
        /// </summary>
        /// <returns>An object of Drawing class.</returns>
        internal Drawing Drawing
        {
            get
            {
                return this.drawing;
            }
        }

        /// <summary>
        /// This method will enable you to get the horizontal dimensions of the image.
        /// </summary>
        /// <returns>An integer value.</returns>
        public long GetExtentCx()
        {
            var inline = this.drawing?.Inline;
            var extent = inline?.Extent;
            return extent?.Cx ?? 0;
        }

        /// <summary>
        /// This method allows you to obtain the vertical dimensions of the image.
        /// </summary>
        /// <returns>An integer value.</returns>
        public long GetExtentCy()
        {
            var inline = this.drawing?.Inline;
            var extent = inline?.Extent;
            return extent?.Cy ?? 0;
        }

        /// <summary>
        /// Invoke this method to extract the collection of images from a Word document.
        /// </summary>
        /// <param name="document">An object of the Document class.</param>
        /// <returns>A collection of images.</returns>
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

        /// <summary>
        /// This method is used to set the border of the image.
        /// </summary>
        /// <param name="borderColor">String value that represents the border color.</param>
        /// <param name="borderWidth">An integer value that represents the border width.</param>
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

        /// <summary>
        /// This method is used to rotate an image.
        /// </summary>
        /// <param name="rotationAngle">An integer value that represents a new angle of the image.</param>
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

        /// <summary>
        /// Call this method to resize the image.
        /// </summary>
        /// <param name="newWidth">An integer value that represents the new width of the image.</param>
        /// <param name="newHeight">An integer value that represents the new height of the image.</param>
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

