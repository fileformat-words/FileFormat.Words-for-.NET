using System;
using System.IO;
using System.IO.Packaging;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FileFormat.Words.Properties;

namespace FileFormat.Words
{
    /// <summary>
    /// This class represents a Word document.
    /// </summary>
    public class Document : IDisposable
    {
        /// <value>
        /// An object of the Parent Document class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDocument;
        private MemoryStream ms;
        private bool disposedValue;

        /// <summary>
        /// Instantiate a new instance of the Document class.
        /// </summary>
        public Document() //Creates a blank WordprocessingML document
        {
            this.ms = new MemoryStream();
            this.wordDocument =
            WordprocessingDocument.Create(this.ms, WordprocessingDocumentType.Document, true);
            this.wordDocument.AddMainDocumentPart();
            this.wordDocument.MainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            this.wordDocument.MainDocumentPart.Document.Body = new DocumentFormat.OpenXml.Wordprocessing.Body();
        }

        /// <summary>
        /// Create an object of the Document class that opens an existing Word document from a file.
        /// </summary>
        /// <param name="docName">String value that represents the document name.</param>
        public Document(string docName) //Opens a WordprocessingML document from a given location
        {

            this.ms = new MemoryStream();
            using (FileStream fs = new FileStream(docName, FileMode.Open))
            {
                fs.CopyTo(this.ms);
            }
            this.wordDocument = WordprocessingDocument.Open(this.ms, true);
        }

        /// <summary>
        /// Create an object of the Document class that opens an existing Word document from a stream.
        /// </summary>
        /// <param name="docStream">An object of the Stream class.</param>
        public Document(Stream docStream) //Opens a WordprocessingML document from stream
        {
            this.ms = new MemoryStream();
            docStream.CopyTo(this.ms);
            this.wordDocument = WordprocessingDocument.Open(this.ms, true);
        }

        /// <summary>
        /// It returns custom built-in document properties.
        /// </summary>
        /// <returns>An object of document properties.</returns>
        public BuiltInDocumentProperties BuiltinDocumentProperties
        {
            get
            {
                BuiltInDocumentProperties prop = new BuiltInDocumentProperties();
                using (var package = Package.Open(this.ms))
                {
                    prop.Author = package.PackageProperties.Creator;
                    prop.CreatedDate = (DateTime)package.PackageProperties.Created;
                    prop.ModifiedBy = package.PackageProperties.LastModifiedBy;
                    prop.ModifiedDate = (DateTime)package.PackageProperties.Modified;
                }
                return prop;
            }
        }

        /// <summary>
        /// Invoke this method to save the document to a file. 
        /// </summary>
        /// <param name="docName">string value represents the document name.</param>
        public void Save(string docName) //Saves the WordprocessingML  document to a given location
        {
            this.wordDocument.Clone(docName);
        }

        /// <summary>
        /// Invoke this method to save the document to a stream. 
        /// </summary>
        /// <param name="docStream">An object of the Stream class.</param>
        public void Save(Stream docStream) //Saves the WordprocessingML document to Stream
        {
            this.wordDocument.Clone(docStream);
        }

        /// <summary>
        /// This method releases unmanaged resources. 
        /// </summary>
        /// <param name="disposing">A boolean value.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                    this.wordDocument.Dispose();
                    this.ms.Dispose();
                }


                disposedValue = true;
            }
        }

        /// <summary>
        /// This method releases unmanaged resources. 
        /// </summary>
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }

}