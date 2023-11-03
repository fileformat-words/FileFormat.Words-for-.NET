using System;
using System.Collections.Generic;
using DF = DocumentFormat.OpenXml;
using PKG = DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using FF = FileFormat.Words.IElements;
using FileFormat.Words;
using System.Linq;

namespace OpenXML.Words.Data
{
    internal static class OOXMLDocData
    {
        private static PKG.WordprocessingDocument staticDoc;
        private static Document ooxmlDoc;// = new Document();
        private static readonly object lockObject = new object(); // Lock object

        internal static void SetPKG(PKG.WordprocessingDocument doc)
        {
            lock (lockObject)
            {
                try
                {
                    staticDoc = doc;
                    ooxmlDoc = new Document();
                }
                catch(Exception ex)
                {
                    string errorMessage = ConstructMessage(ex, "Set OOXML Package");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal static string ConstructMessage(Exception Ex, string Operation)
        {
            return $"Error in Operation {Operation} at OpenXML.Words: {Ex.Message} \n Inner Exception: {Ex.InnerException?.Message ?? "N/A"}";
        }

        internal static void Insert(FF.IElement newElement, int position)
        {
            lock (lockObject)
            {
                List<DF.OpenXmlElement> originalElements =
                    new List<DF.OpenXmlElement>
                    (staticDoc.MainDocumentPart.Document.Body.Elements().ToList());

                try
                {
                    if (newElement is FF.Paragraph ffPara)
                    {
                        WP.Paragraph wpPara = ooxmlDoc.CreateParagraph(ffPara);
                        var elements = staticDoc.MainDocumentPart.Document.Body.Elements();
                        elements.ElementAt(position).InsertBeforeSelf(wpPara);
                    }
                    else if (newElement is FF.Table ffTable)
                    {
                        WP.Table wpTable = ooxmlDoc.CreateTable(ffTable);
                        var elements = staticDoc.MainDocumentPart.Document.Body.Elements();
                        elements.ElementAt(position).InsertBeforeSelf(wpTable);
                    }
                    else if (newElement is FF.Image ffImage)
                    {
                        WP.Paragraph wpImage = ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                        var elements = staticDoc.MainDocumentPart.Document.Body.Elements();
                        elements.ElementAt(position).InsertBeforeSelf(wpImage);
                    }
                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    string errorMessage = ConstructMessage(ex, "Remove OOXML Element(s)"); 
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal static void Update(FF.IElement newElement, int position)
        {
            lock (lockObject) // Lock for thread safety
            {
                List<DF.OpenXmlElement> originalElements =
                    new List<DF.OpenXmlElement>
                    (staticDoc.MainDocumentPart.Document.Body.Elements().ToList());
                try
                {
                    if (newElement is FF.Paragraph ffPara)
                    {
                        WP.Paragraph wpPara = ooxmlDoc.CreateParagraph(ffPara);
                        var elements = staticDoc.MainDocumentPart.Document.Body.Elements();

                        if (position >= 0) //&& position < elements.Count)
                        {
                            var existingElement = elements.ElementAt(position);
                            existingElement.Remove(); // Remove the existing paragraph
                            elements.ElementAt(position).InsertBeforeSelf(wpPara); // Insert the updated paragraph
                        }
                    }
                    else if (newElement is FF.Table ffTable)
                    {
                        WP.Table wpTable = ooxmlDoc.CreateTable(ffTable);
                        var elements = staticDoc.MainDocumentPart.Document.Body.Elements();

                        if (position >= 0)// && position < elements.Count)
                        {
                            var existingElement = elements.ElementAt(position);
                            existingElement.Remove();
                            elements.ElementAt(position).InsertBeforeSelf(wpTable);
                        }
                    }
                    else if (newElement is FF.Image ffImage)
                    {
                        WP.Paragraph wpImage = ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                        var elements = staticDoc.MainDocumentPart.Document.Body.Elements();

                        if (position >= 0) //&& position < elements.Count)
                        {
                            var existingElement = elements.ElementAt(position);
                            existingElement.Remove();
                            elements.ElementAt(position).InsertBeforeSelf(wpImage);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);

                    string errorMessage = ConstructMessage(ex, "Update OOXML Element(s)");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal static void Remove(int position)
        {
            lock (lockObject) // Lock for thread safety
            {
                List<DF.OpenXmlElement> originalElements =
                    new List<DF.OpenXmlElement>
                    (staticDoc.MainDocumentPart.Document.Body.Elements().ToList());
                try
                {
                    var elements = staticDoc.MainDocumentPart.Document.Body.Elements();
                    elements.ElementAt(position).Remove();
                }
                catch(Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);

                    string errorMessage = ConstructMessage(ex, "Remove OOXML Element(s)");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal static void Append(FF.IElement newElement)
        {
            lock (lockObject) // Lock for thread safety
            {
                List<DF.OpenXmlElement> originalElements =
                    new List<DF.OpenXmlElement>
                    (staticDoc.MainDocumentPart.Document.Body.Elements().ToList());
                try
                {
                    //Document ooxmlDoc = new Document();
                    if (newElement is FF.Paragraph ffPara)
                    {
                        WP.Paragraph wpPara = ooxmlDoc.CreateParagraph(ffPara);
                        var sectionPropertiesList = staticDoc.MainDocumentPart.Document.Body.Elements<WP.SectionProperties>().ToList();//staticDoc.MainDocumentPart.Document.Body.Append(wpPara);
                        if (sectionPropertiesList.Any())
                        {
                            // Select the last SectionProperties element, which represents the last section.
                            var lastSectionProperties = sectionPropertiesList.Last();
                            staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpPara, lastSectionProperties);
                            //return lastSectionProperties;
                        }
                        else
                        {
                            staticDoc.MainDocumentPart.Document.Body.Append(wpPara);
                        }
                    }
                    else if (newElement is FF.Table ffTable)
                    {
                        WP.Table wpTable = ooxmlDoc.CreateTable(ffTable);
                        var sectionPropertiesList = staticDoc.MainDocumentPart.Document.Body.Elements<WP.SectionProperties>().ToList();//staticDoc.MainDocumentPart.Document.Body.Append(wpPara);
                        if (sectionPropertiesList.Any())
                        {
                            // Select the last SectionProperties element, which represents the last section.
                            var lastSectionProperties = sectionPropertiesList.Last();
                            staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpTable, lastSectionProperties);
                            //return lastSectionProperties;
                        }
                        else
                        {
                            staticDoc.MainDocumentPart.Document.Body.Append(wpTable);
                        }
                    }
                    else if (newElement is FF.Image ffImage)
                    {
                        WP.Paragraph wpImage = ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                        var sectionPropertiesList = staticDoc.MainDocumentPart.Document.Body.Elements<WP.SectionProperties>().ToList();//staticDoc.MainDocumentPart.Document.Body.Append(wpPara);
                        if (sectionPropertiesList.Any())
                        {
                            // Select the last SectionProperties element, which represents the last section.
                            var lastSectionProperties = sectionPropertiesList.Last();
                            staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpImage, lastSectionProperties);
                            //return lastSectionProperties;
                        }
                        else
                        {
                            staticDoc.MainDocumentPart.Document.Body.Append(wpImage);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    string errorMessage = ConstructMessage(ex, "Append OOXML Element(s)");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal static void Save(System.IO.Stream stream)
        {
            lock (lockObject) // Lock for thread safety
            {
                try
                {
                    //ooxmlDoc.CreateProperties(staticDoc);
                    staticDoc.Clone(stream);
                    staticDoc.Dispose();
                    ooxmlDoc = null;
                }
                catch (Exception ex)
                {
                    string errorMessage = ConstructMessage(ex, "Save OOXML Document");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
    }

}
