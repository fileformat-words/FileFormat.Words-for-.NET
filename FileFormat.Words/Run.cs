using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Underline = DocumentFormat.OpenXml.Wordprocessing.Underline;


namespace FileFormat.Words
{
    public class Run
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.Run wordDocumentRun;
        private string RunText;
        private bool IsBold = false;
        private bool IsItalic = false;

        public Run()
        {
            this.wordDocumentRun = new DocumentFormat.OpenXml.Wordprocessing.Run();
        }
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
        public void AppendChild(Image image)
        {
            this.wordDocumentRun.AppendChild(image.Drawing);
        }

        //public void AppendChild(Graphic graphic)
        //{
        //    this.wordDocumentRun.AppendChild(graphic);
        //}

    }
}

