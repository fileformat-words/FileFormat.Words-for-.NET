using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FileFormat.Words.Properties
{
    /// <summary>
    /// This class represents a collection of built-in document properties.
    /// </summary>
    public class BuiltInDocumentProperties
    {

        /// <summary>
        /// This property is used to get/set the Author name of a Word document.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public String Author { get; set; }

        /// <summary>
        /// This property is used to get/set the subject.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public String Subject { get; set; }

        /// <summary>
        /// This property is used to get/set the title.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public String Title { get; set; }

        /// <summary>
        /// This property is used to get/set the creator name.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public String Creator { get; set; }

        /// <summary>
        /// This property is used to get/set the description.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public String Description { get; set; }

        /// <summary>
        /// This property is used to get/set the keywords.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public String Keywords { get; set; }

        /// <summary>
        /// This property is used to get/set the LastModifiedBy value;
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public String LastModifiedBy { get; set; }

        /// <summary>
        /// This property is used to get/set the CreatedDate value.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// This property is used to get/set the ModifiedDate value.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public DateTime ModifiedDate { get; set; }

        /// <summary>
        /// This property is used to get/set the ModifiedBy value.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public String ModifiedBy { get; set; }

    }
}

