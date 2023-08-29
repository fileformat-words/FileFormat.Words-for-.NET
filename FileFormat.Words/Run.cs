using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;


namespace FileFormat.Words
{
    /// <summary>
    /// The Run class represents a run of characters in a Word document.
    /// </summary>
    public class Run
    {
        /// <value>
        /// An object of the Parent Run class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.Run wordDocumentRun;
        private string RunText;
        private bool IsBold = false;
        private bool IsItalic = false;

        /// <summary>
        /// Instantiate an object of the Run class.
        /// </summary>
        public Run()
        {
            this.wordDocumentRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
        }

        /// <summary>
        /// This property is used to get/set the text of the run.
        /// </summary>
        /// <returns>Returns string value.</returns>
        public string Text
        {
            get
            {
                return this.wordDocumentRun.InnerText;
            }
            set
            {
                this.wordDocumentRun.AppendChild(new Text(value));
            }
        }

        /// <summary>
        /// This property is used to make the run text Bold.
        /// </summary>
        /// <returns>Returns a boolean value.</returns>
        public bool Bold
        {
            get
            {
                return (wordDocumentRun.RunProperties != null && wordDocumentRun.RunProperties.Bold != null);
            }
            set
            {
                if (wordDocumentRun.RunProperties == null)
                {
                    wordDocumentRun.RunProperties = new RunProperties();
                }
                if (value)
                {
                    wordDocumentRun.RunProperties.Bold = new Bold();
                }
                else
                {
                    wordDocumentRun.RunProperties.Bold = null;
                }
            }
        }

        /// <summary>
        /// This property is used to make the run text Italic.
        /// </summary>
        /// <returns>Returns a boolean value.</returns>
        public bool Italic
        {
            get
            {
                return (wordDocumentRun.RunProperties != null && wordDocumentRun.RunProperties.Italic != null);
            }
            set
            {
                if (wordDocumentRun.RunProperties == null)
                {
                    wordDocumentRun.RunProperties = new RunProperties();
                }
                if (value)
                {
                    wordDocumentRun.RunProperties.Italic = new Italic();
                }
                else
                {
                    wordDocumentRun.RunProperties.Italic = null;
                }
            }
        }

        /// <summary>
        /// This property is used to underline the run text.
        /// </summary>
        /// <returns>Returns a boolean value.</returns>
        public bool Underline
        {
            get
            {
                return (wordDocumentRun.RunProperties != null && wordDocumentRun.RunProperties.Underline != null);
            }
            set
            {
                if (wordDocumentRun.RunProperties == null)
                {
                    wordDocumentRun.RunProperties = new RunProperties();
                }
                if (value)
                {
                    wordDocumentRun.RunProperties.Underline = new Underline();
                    wordDocumentRun.RunProperties.Underline.Val = UnderlineValues.Single;
                }
                else
                {
                    wordDocumentRun.RunProperties.Underline = null;
                }
            }
        }

        /// <summary>
        /// This property is used to get/set the font of the run text.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public string FontFamily
        {
            get
            {
                return (wordDocumentRun.RunProperties != null && wordDocumentRun.RunProperties.RunFonts != null) ? wordDocumentRun.RunProperties.RunFonts.Ascii : null;
            }
            set
            {
                if (wordDocumentRun.RunProperties == null)
                {
                    wordDocumentRun.RunProperties = new RunProperties();
                }
                if (wordDocumentRun.RunProperties.RunFonts == null)
                {
                    wordDocumentRun.RunProperties.RunFonts = new RunFonts();
                }
                wordDocumentRun.RunProperties.RunFonts.Ascii = value;
            }
        }

        /// <summary>
        /// This property is used to get/set the font size of the run text.
        /// </summary>
        /// <returns>Returns an integer value.</returns>
        public int FontSize
        {
            get
            {
                return (wordDocumentRun.RunProperties != null && wordDocumentRun.RunProperties.FontSize != null) ? int.Parse(wordDocumentRun.RunProperties.FontSize.Val) : 0;
            }
            set
            {
                if (wordDocumentRun.RunProperties == null)
                {
                    wordDocumentRun.RunProperties = new RunProperties();
                }
                if (wordDocumentRun.RunProperties.FontSize == null)
                {
                    wordDocumentRun.RunProperties.FontSize = new FontSize();
                }
                wordDocumentRun.RunProperties.FontSize.Val = new DocumentFormat.OpenXml.StringValue(value.ToString());
            }
        }

        /// <summary>
        /// This property is used to get/set the text color.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public string Color
        {
            get
            {
                return (wordDocumentRun.RunProperties != null && wordDocumentRun.RunProperties.Color != null) ? wordDocumentRun.RunProperties.Color.Val : null;
            }
            set
            {
                if (wordDocumentRun.RunProperties == null)
                {
                    wordDocumentRun.RunProperties = new RunProperties();
                }
                if (value != null)
                {
                    wordDocumentRun.RunProperties.Color = new Color() { Val = value };
                }
                else
                {
                    wordDocumentRun.RunProperties.Color = null;
                }
            }
        }
        /// <summary>
        /// Call this method to append an object of the Image class.
        /// </summary>
        /// <param name="image">An object of the Image class.</param>
        public void AppendChild(Image image)
        {
            this.wordDocumentRun.AppendChild(image.Drawing);
        }

    }
}

