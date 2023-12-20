using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using DF = DocumentFormat.OpenXml;
using PKG = DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;
using FF = FileFormat.Words.IElements;
using FileFormat.Words;
using System.Linq;

namespace OpenXML.Words.Data
{
    internal class OoxmlDocData
    {
        //private static ConcurrentBag<PKG.WordprocessingDocument> _staticDocBag = new ConcurrentBag<PKG.WordprocessingDocument>();
        //private static readonly List<PKG.WordprocessingDocument> _staticDocs = new List<PKG.WordprocessingDocument>();
        private static ConcurrentDictionary<int, PKG.WordprocessingDocument> _staticDocDict =
            new ConcurrentDictionary<int, PKG.WordprocessingDocument>();
        private static int _staticDocCount = 0;
        private OwDocument _ooxmlDoc;
        private readonly object _lockObject = new object();

        private OoxmlDocData(PKG.WordprocessingDocument doc)
        {
            lock (_lockObject)
            {
                _ooxmlDoc = OwDocument.CreateInstance();
                //_staticDocBag.Add(doc);
                //_staticDocs.Add(doc);
                _staticDocCount++;
                _staticDocDict.TryAdd(_staticDocCount,doc);
            }
        }

        private OoxmlDocData()
        {
            lock (_lockObject)
            {
                _ooxmlDoc = OwDocument.CreateInstance();
            }
        }

        internal static OoxmlDocData CreateInstance(PKG.WordprocessingDocument doc)
        {
            return new OoxmlDocData(doc);
        }
        internal static OoxmlDocData CreateInstance()
        {
            return new OoxmlDocData();
        }


        internal static string ConstructMessage(Exception ex, string operation)
        {
            return $"Error in operation {operation} at OpenXML.Words.Data : {ex.Message} \n Inner Exception: {ex.InnerException?.Message ?? "N/A"}";
        }

        internal void Insert(FF.IElement newElement, int position,Document doc)
        {
            lock (_lockObject)
            {
                /**
                if (!_staticDocBag.TryPeek(out PKG.WordprocessingDocument staticDoc))
                {
                    throw new FileFormatException("No Package is available", new InvalidOperationException());
                }
                **/
                
                //Console.WriteLine("sss : " + doc.GetInstanceInfo());
                //var staticDoc = _staticDocs[instance];

                //_staticDocDict.TryGetValue(instance,out PKG.WordprocessingDocument staticDoc);
                _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out PKG.WordprocessingDocument staticDoc);

                if (staticDoc?.MainDocumentPart?.Document?.Body == null) throw new FileFormatException("Package or Document or Body is null",new NullReferenceException());
                
                var enumerable = staticDoc.MainDocumentPart.Document.Body.Elements().ToList();
                var originalElements = new List<DF.OpenXmlElement>(enumerable);

                var elements = staticDoc.MainDocumentPart.Document.Body.Elements();

                try
                {
                    switch (newElement)
                    {
                        case FF.Paragraph ffPara:
                            var wpPara = _ooxmlDoc.CreateParagraph(ffPara);
                            elements.ElementAt(position).InsertBeforeSelf(wpPara);
                            break;

                        case FF.Table ffTable:
                            var wpTable = _ooxmlDoc.CreateTable(ffTable);
                            elements.ElementAt(position).InsertBeforeSelf(wpTable);
                            break;

                        case FF.Image ffImage:
                            var wpImage = _ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                            elements.ElementAt(position).InsertBeforeSelf(wpImage);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    var errorMessage = ConstructMessage(ex, "Remove OOXML Element(s)");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal void Update(FF.IElement newElement, int position,Document doc)
        {
            lock (_lockObject) 
            {
                /**
                if (!_staticDocBag.TryPeek(out PKG.WordprocessingDocument staticDoc))
                {
                    throw new FileFormatException("No Package is available", new InvalidOperationException());
                }**/

                //var staticDoc = _staticDocs[instance];

                //_staticDocDict.TryGetValue(instance, out PKG.WordprocessingDocument staticDoc);
                _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out PKG.WordprocessingDocument staticDoc);

                if (staticDoc?.MainDocumentPart?.Document?.Body == null) throw new FileFormatException("Package or Document or Body is null", new NullReferenceException());
                
                var enumerable = staticDoc.MainDocumentPart.Document.Body.Elements().ToList();
                var originalElements = new List<DF.OpenXmlElement>(enumerable);
                
                try
                {
                    if (position >= 0)
                    {
                        var elements = staticDoc.MainDocumentPart.Document.Body.Elements();
                        elements.ElementAt(position).Remove();
                        var enumerable1 = elements.ToList();
                        var existingElement = enumerable1.ElementAt(position);
                        switch (newElement)
                        {
                            case FF.Paragraph ffPara:
                                var wpPara = _ooxmlDoc.CreateParagraph(ffPara);
                                enumerable1.ElementAt(position).InsertBeforeSelf(wpPara);
                                break;
                            case FF.Table ffTable:
                                var wpTable = _ooxmlDoc.CreateTable(ffTable);
                                enumerable1.ElementAt(position).InsertBeforeSelf(wpTable);
                                break;
                            case FF.Image ffImage:
                                var wpImage = _ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                                enumerable1.ElementAt(position).InsertBeforeSelf(wpImage);
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    var errorMessage = ConstructMessage(ex, "Update OOXML Element(s)");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal void Remove(int position,Document doc)
        {
            lock (_lockObject) 
            {
                /**
                if (!_staticDocBag.TryPeek(out PKG.WordprocessingDocument staticDoc))
                {
                    throw new FileFormatException("No Package is available", new InvalidOperationException());
                }**/


                //var staticDoc = _staticDocs[instance];

                //_staticDocDict.TryGetValue(instance, out PKG.WordprocessingDocument staticDoc);
                _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out PKG.WordprocessingDocument staticDoc);

                if (staticDoc?.MainDocumentPart?.Document?.Body == null) throw new FileFormatException("Package or Document or Body is null", new NullReferenceException());

                var enumerable = staticDoc.MainDocumentPart.Document.Body.Elements().ToList();
                var originalElements = new List<DF.OpenXmlElement>(enumerable);

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
                    var errorMessage = ConstructMessage(ex, "Remove OOXML Element(s)");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal void Append(FF.IElement newElement,Document doc)
        {
            lock (_lockObject) 
            {
                /**
                if (!_staticDocBag.TryPeek(out PKG.WordprocessingDocument staticDoc))
                {
                    throw new FileFormatException("No Package is available", new InvalidOperationException());
                }**/

                //var staticDoc = _staticDocs[instance];

                //_staticDocDict.TryGetValue(instance, out PKG.WordprocessingDocument staticDoc);

                _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out PKG.WordprocessingDocument staticDoc);

                if (staticDoc?.MainDocumentPart?.Document?.Body == null) throw new FileFormatException("Package or Document or Body is null", new NullReferenceException());

                var enumerable = staticDoc.MainDocumentPart.Document.Body.Elements().ToList();
                var originalElements = new List<DF.OpenXmlElement>(enumerable);

                var sectionPropertiesList = staticDoc.MainDocumentPart.Document.Body.Elements<WP.SectionProperties>().ToList();
                WP.SectionProperties lastSectionProperties = null;
                if (sectionPropertiesList.Any()) lastSectionProperties = sectionPropertiesList.Last();

                try
                {
                    switch (newElement)
                    {
                        case FF.Paragraph ffPara:
                            var wpPara = _ooxmlDoc.CreateParagraph(ffPara);
                            if (lastSectionProperties != null) staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpPara, lastSectionProperties);
                            else staticDoc.MainDocumentPart.Document.Body.Append(wpPara);
                            break;
                        case FF.Table ffTable:
                            var wpTable = _ooxmlDoc.CreateTable(ffTable);
                            if (lastSectionProperties != null) staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpTable, lastSectionProperties);
                            else staticDoc.MainDocumentPart.Document.Body.Append(wpTable);
                            break;
                        case FF.Image ffImage:
                            var wpImage = _ooxmlDoc.CreateImage(ffImage, staticDoc.MainDocumentPart);
                            if (lastSectionProperties != null) staticDoc.MainDocumentPart.Document.Body.InsertBefore(wpImage, lastSectionProperties);
                            else staticDoc.MainDocumentPart.Document.Body.Append(wpImage);
                            break;
                    }

                }
                catch (Exception ex)
                {
                    // Rollback changes by restoring the original elements
                    staticDoc.MainDocumentPart.Document.Body.RemoveAllChildren();
                    staticDoc.MainDocumentPart.Document.Body.Append(originalElements);
                    var errorMessage = ConstructMessage(ex, "Append OOXML Element(s)");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        internal void Save(System.IO.Stream stream, Document doc)
        {
            lock (_lockObject) 
            {
                try
                {
                    /**
                    if (!_staticDocBag.TryPeek(out PKG.WordprocessingDocument staticDoc))
                    {
                        throw new FileFormatException("No Package is available", new InvalidOperationException());
                    }**/


                    //var staticDoc = _staticDocs[instance];

                    //_staticDocDict.TryGetValue(instance, out PKG.WordprocessingDocument staticDoc);

                    _staticDocDict.TryGetValue(doc.GetInstanceInfo(), out PKG.WordprocessingDocument staticDoc);

                    //_ooxmlDoc.CreateProperties(_staticDoc);
                    staticDoc.Clone(stream);
                    //_staticDoc.Dispose();
                    //_ooxmlDoc = null;
                }
                catch (Exception ex)
                {
                    var errorMessage = ConstructMessage(ex, "Save OOXML OWDocument");
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
                /**
                if (_staticDocBag != null)
                {
                    while (_staticDocBag.TryTake(out var staticDoc))
                    {
                        if (staticDoc == null) continue;
                        staticDoc.Dispose();
                    }
                }
                **/

                if (_ooxmlDoc == null) return;
                _ooxmlDoc.Dispose();
                _ooxmlDoc = null;
            }
            // Dispose of unmanaged resources
        }
    }
}
