using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FileFormat.Words.Properties
{
    public class BuiltInDocumentProperties
    {

        public String Author { get; set; }

        public String Subject { get; set; }

        public String Title { get; set; }

        public String Creator { get; set; }

        public String Description { get; set; }

        public String Keywords { get; set; }

        public String LastModifiedBy { get; set; }

        public DateTime CreatedDate { get; set; }

        public DateTime ModifiedDate { get; set; }

        public String ModifiedBy { get; set; }

    }
}

