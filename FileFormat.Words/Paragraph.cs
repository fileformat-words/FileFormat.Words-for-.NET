using System;
using System.IO.Packaging;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace FileFormat.Words
{
    public class Paragraph
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.Paragraph wordDocumentParagraph;
        private string ParaText;
        private Justification justification;

        public Paragraph()
        {
            this.wordDocumentParagraph = new DocumentFormat.OpenXml.Wordprocessing.Paragraph();
        }

        public string Text
        {
            get
            {
                return wordDocumentParagraph.InnerText;
            }
            set
            {
                DocumentFormat.OpenXml.Wordprocessing.Run run = this.wordDocumentParagraph.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
                run.AppendChild(new Text(value));
            }
        }

        public void AppendChild(Run run)
        {
            this.wordDocumentParagraph.AppendChild(run.wordDocumentRun);
        }



        public IEnumerable<Run> GetRuns()
        {
            var runs = this.wordDocumentParagraph.
                Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>();
            Run run1;
            List<Run> lst = new List<Run>();
            foreach (DocumentFormat.OpenXml.Wordprocessing.Run run in runs.ToArray())
            {
                run1 = new Run();
                run1.Text = run.InnerText;
                run1.Bold = (run.RunProperties != null && run.RunProperties.Bold != null);
                run1.Italic = (run.RunProperties != null && run.RunProperties.Italic != null);
                run1.Underline = (run.RunProperties != null && run.RunProperties.Underline != null);
                run1.FontFamily = (run.RunProperties != null && run.RunProperties.RunFonts != null) ? run.RunProperties.RunFonts.Ascii : null;
                run1.FontSize = (run.RunProperties != null && run.RunProperties.FontSize != null) ? int.Parse(run.RunProperties.FontSize.Val) : 0;
                run1.Color = (run.RunProperties != null && run.RunProperties.Color != null) ? run.RunProperties.Color.Val : null;

                lst.Add(run1);
            }
            return lst;
        }


        public string Style
        {
            get
            {
                if (this.wordDocumentParagraph.ParagraphProperties != null
                    && this.wordDocumentParagraph.ParagraphProperties.ParagraphStyleId != null)
                {
                    return this.wordDocumentParagraph.ParagraphProperties.ParagraphStyleId.Val;
                }

                return null;
            }
            set
            {
                Console.WriteLine("Value = " + value);
                if (this.wordDocumentParagraph.ParagraphProperties == null)
                {
                    this.wordDocumentParagraph.ParagraphProperties = new ParagraphProperties();
                }

                this.wordDocumentParagraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = value };
            }
        }
        public string LinesSpacing
        {
            get
            {
                if (this.wordDocumentParagraph.ParagraphProperties != null)
                {
                    return this.wordDocumentParagraph.ParagraphProperties.SpacingBetweenLines.Line.Value;
                }
                return null;
            }
            set
            {
                if (this.wordDocumentParagraph.ParagraphProperties == null)
                {
                    this.wordDocumentParagraph.ParagraphProperties = new ParagraphProperties();
                }
                SpacingBetweenLines spacing = new SpacingBetweenLines() { Line = value };
                this.wordDocumentParagraph.ParagraphProperties.Append(spacing);

            }
        }

        public string Align
        {
            get
            {
                if (this.wordDocumentParagraph.ParagraphProperties != null && this.wordDocumentParagraph.ParagraphProperties.Justification != null)
                {
                    return this.wordDocumentParagraph.ParagraphProperties.Justification.Val;
                }
                return null;
            }
            set
            {
                if (this.wordDocumentParagraph.ParagraphProperties == null)
                {
                    this.wordDocumentParagraph.ParagraphProperties = new ParagraphProperties();
                }

                switch (value)
                {

                    case "Center":
                        this.justification = new Justification() { Val = JustificationValues.Center };
                        this.wordDocumentParagraph.ParagraphProperties.Append(this.justification);
                        break;
                    case "Right":
                        this.justification = new Justification() { Val = JustificationValues.Right };
                        this.wordDocumentParagraph.ParagraphProperties.Append(this.justification);
                        break;
                    case "Left":
                        this.justification = new Justification() { Val = JustificationValues.Left };
                        this.wordDocumentParagraph.ParagraphProperties.Append(this.justification);
                        break;
                    case "Both":
                        this.justification = new Justification() { Val = JustificationValues.Both };
                        this.wordDocumentParagraph.ParagraphProperties.Append(this.justification);
                        break;

                }

            }
        }
        public string Indent
        {
            get
            {
                if (this.wordDocumentParagraph.ParagraphProperties != null && this.wordDocumentParagraph.ParagraphProperties.Indentation.Start != null)
                {
                    return this.wordDocumentParagraph.ParagraphProperties.Indentation.Start.Value;
                }
                return null;
            }
            set
            {
                if (this.wordDocumentParagraph.ParagraphProperties == null)
                {
                    this.wordDocumentParagraph.ParagraphProperties = new ParagraphProperties();
                }
                Indentation indent = new Indentation() { Start = value };
                this.wordDocumentParagraph.ParagraphProperties.Append(indent);
            }
        }
        public string FirstLineIndent
        {
            get
            {
                if (this.wordDocumentParagraph.ParagraphProperties != null && this.wordDocumentParagraph.ParagraphProperties.Indentation != null)
                {
                    return this.wordDocumentParagraph.ParagraphProperties.Indentation.FirstLine.Value;
                }
                return null;
            }
            set
            {
                if (this.wordDocumentParagraph.ParagraphProperties == null)
                {
                    this.wordDocumentParagraph.ParagraphProperties = new ParagraphProperties();
                }
                Indentation indent = new Indentation() { FirstLine = value };
                this.wordDocumentParagraph.ParagraphProperties.Append(indent);
            }
        }
        public string LeftIndent
        {
            get
            {
                if (this.wordDocumentParagraph.ParagraphProperties != null)
                {
                    return this.wordDocumentParagraph.ParagraphProperties.Indentation.Left.Value;
                }
                return null;
            }
            set
            {
                if (this.wordDocumentParagraph.ParagraphProperties == null)
                {
                    this.wordDocumentParagraph.ParagraphProperties = new ParagraphProperties();
                }
                Indentation indent = new Indentation() { Left = value };
                this.wordDocumentParagraph.ParagraphProperties.Append(indent);
            }
        }
        public string RihgtIndent
        {
            get
            {
                if (this.wordDocumentParagraph.ParagraphProperties != null)
                {
                    return this.wordDocumentParagraph.ParagraphProperties.Indentation.Right.Value;
                }
                return null;
            }
            set
            {
                if (this.wordDocumentParagraph.ParagraphProperties == null)
                {
                    this.wordDocumentParagraph.ParagraphProperties = new ParagraphProperties();
                }
                Indentation indent = new Indentation() { Right = value };
                this.wordDocumentParagraph.ParagraphProperties.Append(indent);
            }
        }

        public bool IsHeading
        {
            get
            {
                if (this.wordDocumentParagraph.ParagraphProperties != null
                    && this.wordDocumentParagraph.ParagraphProperties.ParagraphStyleId != null)
                {
                    string styleId = this.wordDocumentParagraph.ParagraphProperties.ParagraphStyleId.Val;

                    if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                return false;
            }
        }

    }
}

