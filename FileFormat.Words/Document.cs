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
    public class Document : IDisposable
    {
        protected internal DocumentFormat.OpenXml.Packaging.WordprocessingDocument wordDocument;
        private MemoryStream ms;
        private bool disposedValue;


        public Document() //Creates a blank WordprocessingML document
        {
            this.ms = new MemoryStream();
            this.wordDocument =
            WordprocessingDocument.Create(this.ms, WordprocessingDocumentType.Document, true);
            this.wordDocument.AddMainDocumentPart();
            this.wordDocument.MainDocumentPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
            this.wordDocument.MainDocumentPart.Document.Body = new DocumentFormat.OpenXml.Wordprocessing.Body();
        }

        public Document(string docName) //Opens a WordprocessingML document from a given location
        {

            this.ms = new MemoryStream();
            using (FileStream fs = new FileStream(docName, FileMode.Open))
            {
                fs.CopyTo(this.ms);
            }
            this.wordDocument = WordprocessingDocument.Open(this.ms, true);
        }

        public Document(Stream docStream) //Opens a WordprocessingML document from stream
        {
            this.ms = new MemoryStream();
            docStream.CopyTo(this.ms);
            this.wordDocument = WordprocessingDocument.Open(this.ms, true);
        }

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

        public void Save(string docName) //Saves the WordprocessingML  document to a given location
        {
            this.wordDocument.Clone(docName);
        }

        public void Save(Stream docStream) //Saves the WordprocessingML document to Stream
        {
            this.wordDocument.Clone(docStream);
        }

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


        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }

}