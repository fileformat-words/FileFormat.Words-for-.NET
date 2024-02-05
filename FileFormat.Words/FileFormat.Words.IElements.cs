using System.Collections.Generic;
using System.Linq;
namespace FileFormat.Words.IElements
{
    /// <summary>
    /// Represents an element in a Word document.
    /// </summary>
    public interface IElement
    {
        /// <summary>
        /// Gets the unique identifier of the element.
        /// </summary>
        int ElementId { get; }
    }

    public class Indentation
    {
        public double Left { get; set; }
        public double Right { get; set; }
        public double FirstLine { get; set; }
        public double Hanging { get; set; }
    }

    /// <summary>
    /// Represents a paragraph element in a Word document.
    /// </summary>
    public class Paragraph : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the paragraph.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets or sets the text content of the paragraph.
        /// </summary>
        public string Text { get; private set; }

        /// <summary>
        /// Gets the list of runs (text fragments) within the paragraph.
        /// </summary>
        public List<Run> Runs { get; }

        /// <summary>
        /// Gets or sets the style of the paragraph.
        /// </summary>
        public string Style { get; set; }

        /// <summary>
        /// Gets or Sets Alignment of the word paragraph
        /// </summary>
        public string Alignment { get; set; }

        /// <summary>
        /// Gets or Sets Indentation of the word paragraph
        /// </summary>
        public Indentation Indentation { get; set; }

        /// <summary>
        /// Gets or sets the numbering ID for the paragraph.
        /// </summary>
        public int? NumberingId { get; set; }

        /// <summary>
        /// Gets or sets the numbering level for the paragraph.
        /// </summary>
        public int? NumberingLevel { get; set; }

        /// <summary>
        /// Gets or sets whether the paragraph has bullet points.
        /// </summary>
        public bool IsBullet { get; set; }

        /// <summary>
        /// Gets or sets whether the paragraph has numbering.
        /// </summary>
        public bool IsNumbered { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Paragraph"/> class.
        /// </summary>
        public Paragraph()
        {
            Runs = new List<Run>();
            Style = "Normal";
            Indentation = new Indentation();
            NumberingId = null;
            NumberingLevel = null;
            IsBullet = false;
            IsNumbered = false;
            UpdateText(); // Initialize the Text property
        }

        /// <summary>
        /// Adds a run (text fragment) to the paragraph and sets its parent paragraph.
        /// </summary>
        /// <param name="run">The run to add to the paragraph.</param>
        public void AddRun(Run run)
        {
            run.ParentParagraph = this;
            Runs.Add(run);
            UpdateText(); // Update the Text property when a new run is added
        }

        internal void UpdateText()
        {
            Text = string.Join("", Runs.Select(run => run.Text));
        }
    }

    /// <summary>
    /// Provides predefined heading styles.
    /// </summary>
    public static class Headings
    {
        /// <summary>
        /// Gets the value representing Heading1.
        /// </summary>
        public static string Heading1 { get; } = "Heading1";

        /// <summary>
        /// Gets the value representing Heading2.
        /// </summary>
        public static string Heading2 { get; } = "Heading2";

        /// <summary>
        /// Gets the value representing Heading3.
        /// </summary>
        public static string Heading3 { get; } = "Heading3";

        /// <summary>
        /// Gets the value representing Heading4.
        /// </summary>
        public static string Heading4 { get; } = "Heading4";

        /// <summary>
        /// Gets the value representing Heading5.
        /// </summary>
        public static string Heading5 { get; } = "Heading5";

        /// <summary>
        /// Gets the value representing Heading6.
        /// </summary>
        public static string Heading6 { get; } = "Heading6";

        /// <summary>
        /// Gets the value representing Heading7.
        /// </summary>
        public static string Heading7 { get; } = "Heading7";

        /// <summary>
        /// Gets the value representing Heading8.
        /// </summary>
        public static string Heading8 { get; } = "Heading8";

        /// <summary>
        /// Gets the value representing Heading9.
        /// </summary>
        public static string Heading9 { get; } = "Heading9";
    }


    /// <summary>
    /// Represents a run of text within a paragraph.
    /// </summary>
    public class Run
    {
        private string _text;
        /// <summary>
        /// Gets or sets the text content of the run.
        /// </summary>
        public string Text
        {
            
            get => _text;
            set
            {
                _text = value;
                if (ParentParagraph != null)
                {
                    ParentParagraph.UpdateText();
                }
            }
    }

        /// <summary>
        /// Gets or sets the font family of the run.
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// Gets or sets the font size of the run.
        /// </summary>
        public int FontSize { get; set; }

        /// <summary>
        /// Gets or sets the color of the run's text.
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Gets or sets whether the run's text is bold.
        /// </summary>
        public bool Bold { get; set; }

        /// <summary>
        /// Gets or sets whether the run's text is italic.
        /// </summary>
        public bool Italic { get; set; }

        /// <summary>
        /// Gets or sets whether the run's text is underlined.
        /// </summary>
        public bool Underline { get; set; }

        internal Paragraph ParentParagraph { get; set; }
    }

    /// <summary>
    /// Provides predefined colors with hexadecimal values.
    /// </summary>
    public static class Colors
    {
        /// <summary>
        /// Gets the hexadecimal value for the color Black (000000).
        /// </summary>
        public static string Black { get; } = "000000";

        /// <summary>
        /// Gets the hexadecimal value for the color White (FFFFFF).
        /// </summary>
        public static string White { get; } = "FFFFFF";

        /// <summary>
        /// Gets the hexadecimal value for the color Red (FF0000).
        /// </summary>
        public static string Red { get; } = "FF0000";

        /// <summary>
        /// Gets the hexadecimal value for the color Green (00FF00).
        /// </summary>
        public static string Green { get; } = "00FF00";

        /// <summary>
        /// Gets the hexadecimal value for the color Blue (0000FF).
        /// </summary>
        public static string Blue { get; } = "0000FF";

        /// <summary>
        /// Gets the hexadecimal value for the color Yellow (FFFF00).
        /// </summary>
        public static string Yellow { get; } = "FFFF00";

        /// <summary>
        /// Gets the hexadecimal value for the color Cyan (00FFFF).
        /// </summary>
        public static string Cyan { get; } = "00FFFF";

        /// <summary>
        /// Gets the hexadecimal value for the color Magenta (FF00FF).
        /// </summary>
        public static string Magenta { get; } = "FF00FF";

        /// <summary>
        /// Gets the hexadecimal value for the color Gray (808080).
        /// </summary>
        public static string Gray { get; } = "808080";

        /// <summary>
        /// Gets the hexadecimal value for the color Silver (C0C0C0).
        /// </summary>
        public static string Silver { get; } = "C0C0C0";

        /// <summary>
        /// Gets the hexadecimal value for the color Maroon (800000).
        /// </summary>
        public static string Maroon { get; } = "800000";

        /// <summary>
        /// Gets the hexadecimal value for the color Olive (808000).
        /// </summary>
        public static string Olive { get; } = "808000";

        /// <summary>
        /// Gets the hexadecimal value for the color Green (008000).
        /// </summary>
        public static string Teal { get; } = "008000";

        /// <summary>
        /// Gets the hexadecimal value for the color Navy (000080).
        /// </summary>
        public static string Navy { get; } = "000080";

        /// <summary>
        /// Gets the hexadecimal value for the color Purple (800080).
        /// </summary>
        public static string Purple { get; } = "800080";

        /// <summary>
        /// Gets the hexadecimal value for the color Orange (FFA500).
        /// </summary>
        public static string Orange { get; } = "FFA500";

        /// <summary>
        /// Gets the hexadecimal value for the color Lime (00FF00).
        /// </summary>
        public static string Lime { get; } = "00FF00";

        /// <summary>
        /// Gets the hexadecimal value for the color Aqua (00FFFF).
        /// </summary>
        public static string Aqua { get; } = "00FFFF";

        /// <summary>
        /// Gets the hexadecimal value for the color Fuchsia (FF00FF).
        /// </summary>
        public static string Fuchsia { get; } = "FF00FF";

        /// <summary>
        /// Gets the hexadecimal value for the color Silver (C0C0C0).
        /// </summary>
        public static string LimeGreen { get; } = "32CD32";
    }

    /// <summary>
    /// Represents an image element in a Word document.
    /// </summary>
    public class Image : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the image.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets or sets the binary image data.
        /// </summary>
        public byte[] ImageData { get; set; }

        /// <summary>
        /// Gets or sets the height of the image.
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Gets or sets the width of the image.
        /// </summary>
        public int Width { get; set; }
    }
    /// <summary>
    /// Represents a table element in a Word document.
    /// </summary>
    public class Table : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the table.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets or sets the table style.
        /// </summary>
        public string Style { get; set; }

        /// <summary>
        /// Gets or sets the list of rows within the table.
        /// </summary>
        public List<Row> Rows { get; set; }

        /// <summary>
        /// Gets or sets the column properties of the table.
        /// </summary>
        public Column Column { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Table"/> class with empty rows and default column properties.
        /// </summary>
        public Table()
        {
            Rows = new List<Row>();
            Column = new Column();
            Style = "TableGrid";
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Table"/> class with a specified number of rows and columns.
        /// </summary>
        /// <param name="rows">The number of rows in the table.</param>
        /// <param name="columns">The number of columns in the table.</param>
        public Table(int rows, int columns)
        {
            Rows = new List<Row>();
            Column = new Column();

            for (var i = 0; i < rows; i++)
            {
                var row = new Row();
                row.Cells = new List<Cell>();

                for (var j = 0; j < columns; j++)
                {
                    var cellContent = "";
                    var paragraph = new Paragraph();
                    paragraph.AddRun(new Run { Text = cellContent });

                    var cell = new Cell { Paragraphs = new List<Paragraph> { paragraph } };
                    row.Cells.Add(cell);
                }

                Rows.Add(row);
            }
        }
    }
    /// <summary>
    /// Represents a row within a table in a Word document.
    /// </summary>
    public class Row
    {
        /// <summary>
        /// Gets or sets the list of cells within the row.
        /// </summary>
        public List<Cell> Cells { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Row"/> class with empty cells.
        /// </summary>
        public Row()
        {
            Cells = new List<Cell>();
        }
    }

    /// <summary>
    /// Represents a cell within a row of a table in a Word document.
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Gets or sets the list of paragraphs within the cell.
        /// </summary>
        public List<Paragraph> Paragraphs { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Cell"/> class with empty paragraphs.
        /// </summary>
        public Cell()
        {
            Paragraphs = new List<Paragraph>();
        }
    }

    /// <summary>
    /// Represents column properties of a table in a Word document.
    /// </summary>
    public class Column
    {
        /// <summary>
        /// Gets or sets the width of the column.
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Column"/> class with a default width of 0.
        /// </summary>
        public Column()
        {
            Width = 0;
        }
    }

    /// <summary>
    /// Represents a section element in a Word document.
    /// </summary>
    public class Section : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the section.
        /// </summary>
        public int ElementId { get; internal set; }

        /// <summary>
        /// Gets the page size properties for the section.
        /// </summary>
        public PageSize PageSize { get; internal set; }

        /// <summary>
        /// Gets the page margin properties for the section.
        /// </summary>
        public PageMargin PageMargin { get; internal set; }
        internal Section()
        {
            //Do nothing
        }
    }

    /// <summary>
    /// Represents the page size properties of a section in a Word document.
    /// </summary>
    public class PageSize
    {
        /// <summary>
        /// Gets sets the height of the page.
        /// </summary>
        public int Height { get; internal set; }

        /// <summary>
        /// Gets the width of the page.
        /// </summary>
        public int Width { get; internal set; }

        /// <summary>
        /// Gets the orientation of the page (e.g., "Portrait" or "Landscape").
        /// </summary>
        public string Orientation { get; internal set; }
        internal PageSize()
        {
            //Do nothing
        }
    }

    /// <summary>
    /// Represents the page margin properties of a section in a Word document.
    /// </summary>
    public class PageMargin
    {
        /// <summary>
        /// Gets the top margin of the page.
        /// </summary>
        public int Top { get; internal set; }

        /// <summary>
        /// Gets the right margin of the page.
        /// </summary>
        public int Right { get; internal set; }

        /// <summary>
        /// Gets the bottom margin of the page.
        /// </summary>
        public int Bottom { get; internal set; }

        /// <summary>
        /// Gets the left margin of the page.
        /// </summary>
        public int Left { get; internal set; }

        /// <summary>
        /// Gets the header margin of the page.
        /// </summary>
        public int Header { get; internal set; }

        /// <summary>
        /// Gets the footer margin of the page.
        /// </summary>
        public int Footer { get; internal set; }
        internal PageMargin()
        {
            //Do nothing
        }
    }

    /// <summary>
    /// Represents an unknown element in a Word document.
    /// </summary>
    public class Unknown : IElement
    {
        /// <summary>
        /// Gets the unique identifier of the unknown element.
        /// </summary>
        public int ElementId { get; internal set; }

        internal Unknown()
        {
            // Do nothing
        }
    }
    /// <summary>
    /// Represents Styles associated with different elements.
    /// </summary>
    public class ElementStyles
    {
        /// <summary>
        /// Gets the fonts defined in theme.
        /// </summary>
        public List<string> ThemeFonts { get; internal set; }
        /// <summary>
        /// Gets the fonts defined in FontTable
        /// </summary>
        public List<string> TableFonts { get; internal set; }
        /// <summary>
        /// Gets the Paragraph Styles
        /// </summary>
        public List<string> ParagraphStyles { get; internal set; }
        /// <summary>
        /// Gets the Table Styles
        /// </summary>
        public List<string> TableStyles { get; internal set; }
        /// <summary>
        /// Initializes all Styles.
        /// </summary>
        public ElementStyles()
        {
            ThemeFonts = new List<string>();
            TableFonts = new List<string>();
            ParagraphStyles = new List<string>();
            TableStyles = new List<string>();
        }
    }

}