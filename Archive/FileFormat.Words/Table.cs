using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace FileFormat.Words.Table
{
    /// <summary>
    /// This class represents a table.
    /// </summary>
    public class Table
    {
        /// <value>
        /// An object of the Parent Table class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.Table table;
        bool haveHeaders = false;
        private string numberOfTables;
        private string numberOfRows;
        private string numberOfColumns;
        private string numberOfCells;
        private string cellWidth;
        private string tableBorder;
        private string tablePosition;
        private List<string> tableHeaders = new List<string>();
        private List<string> existingTableHeaders = new List<string>();

        /// <summary>
        /// Instantiate an object of the Table class.
        /// </summary>
        public Table()
        {
            this.table = new DocumentFormat.OpenXml.Wordprocessing.Table();
        }

        /// <summary>
        /// Call this method to set the table header values. 
        /// </summary>
        /// <param name="val">A string value.</param>
        public void TableHeaders(string val)
        {
            tableHeaders.Add(val);
        }

        /// <summary>
        /// Invoke this method to add a row to the table.
        /// </summary>
        /// <param name="tr">An object of the table row.</param>
        public void Append(TableRow tr)
        {

            if (!haveHeaders)
            {
                List<string> cellWidth = new List<string>();

                DocumentFormat.OpenXml.Wordprocessing.TableRow tb = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                if (cellWidth.Count() == 0)
                {
                    for (int i = 1; i < tr.tableRow.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList().Count() + 1; i++)
                    {
                        tableHeaders.Add("column" + i);
                        cellWidth.Add(tr.tableRow.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToArray()[i - 1].TableCellProperties.TableCellWidth.Width);

                    }
                }

                for (int obj = 0; obj < tr.tableRow.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList().Count(); obj++)
                {
                    DocumentFormat.OpenXml.Wordprocessing.TableCell tableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                    DocumentFormat.OpenXml.Wordprocessing.TableCellProperties celp = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties();
                    celp.Append(new DocumentFormat.OpenXml.Wordprocessing.TableCellWidth() { Width = cellWidth[obj] });
                    tableCell.Append(celp);
                    DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run(new Text(tableHeaders[obj]));
                    run.AppendChild(new RunProperties(new Bold()));
                    tableCell.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(run));
                    tb.Append(tableCell);

                }
                this.table.Append(tb);
                this.table.Append(tr.tableRow);
                haveHeaders = true;
                tableHeaders = null;
            }
            else 
                this.table.Append(tr.tableRow);

        }

        /// <summary>
        /// Invoke this method to change the inner text of a specific cell into a table.
        /// </summary>
        /// <param name="path">Path of the Word document.</param>
        /// <param name="tableIndex">Represents the index of a table in a document.</param>
        /// <param name="tableRow">Index of the row.</param>
        /// <param name="tableCell">Cell index.</param>
        /// <param name="txt">A string value that will replace the existing text.</param>
        /// <returns>A string value.</returns>
        public string ChangeTextInCell(string path, int tableIndex, int tableRow, int tableCell, string txt)
        {
            using (WordprocessingDocument doc =
            WordprocessingDocument.Open(path, true))
            {
                if (tableIndex >= doc.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList().Count())
                    return "table index out of range";

                // Find the first table in the document.
                DocumentFormat.OpenXml.Wordprocessing.Table table =
                doc.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList()[tableIndex];

                if (tableRow >= table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ToList().Count())
                    return "table row index out of range";

                // Find the specific row in the table.
                DocumentFormat.OpenXml.Wordprocessing.TableRow row = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ElementAt(tableRow);

                if (tableCell >= row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList().Count())
                    return "table cell index out of range";
                // Find the specific cell in the row.
                DocumentFormat.OpenXml.Wordprocessing.TableCell cell = row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAt(tableCell);

                // Find the first paragraph in the table cell.
                DocumentFormat.OpenXml.Wordprocessing.Paragraph p = cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();

                // Find the first run in the paragraph.
                DocumentFormat.OpenXml.Wordprocessing.Run r = p.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().First();

                // Set the text for the run.
                Text t = r.Elements<Text>().First();
                t.Text = txt;

                return "Cell updated successfully";
            }
        }

        /// <summary>
        /// This property is used to get/set the table headers.
        /// </summary>
        /// <returns>A collection of table headers.</returns>
        public List<string> ExistingTableHeaders
        {
            get { return existingTableHeaders; }
            set
            {
                existingTableHeaders.Add(value.ToString());
            }
        }

        /// <summary>
        /// Call this method to append table properties to the Table's object.
        /// </summary>
        /// <param name="tp">An instance of the TableProperties class.</param>
        public void AppendChild(TableProperties tp)
        {
            this.table.AppendChild(tp.tableProperties);
        }

        /// <summary>
        /// This property returns the total count of table rows.
        /// </summary>
        /// <returns>An integer value.</returns>
        public int getTableRows
        {
            get
            {
                return this.table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ToList().Count;
            }
        }

        /// <summary>
        /// This property returns the total count of table cells.
        /// </summary>
        /// <returns>An integer value.</returns>
        public int getTableCells
        {
            get
            {
                return this.table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList().Count;
            }
        }

        /// <summary>
        /// This property returns the total count of table columns.
        /// </summary>
        /// <returns>An integer value.</returns>
        public int getTableColumns
        {
            get
            {

                return this.table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList().Count;
            }
        }

        /// <summary>
        /// This property is used to set/get the count of total tables.
        /// </summary>
        /// <returns>A string value.</returns>
        public string NumberOfTables
        {
            get { return numberOfTables; }
            set { numberOfTables = value; }
        }

        /// <summary>
        /// This property is used to set/get the count of total table rows.
        /// </summary>
        /// <returns>A string value.</returns>
        public string NumberOfRows
        {
            get { return numberOfRows; }
            set { numberOfRows = value; }
        }

        /// <summary>
        /// This property is used to set/get the count of total table columns.
        /// </summary>
        /// <returns>A string value.</returns>
        public string NumberOfColumns
        {
            get { return numberOfColumns; }
            set { numberOfColumns = value; }
        }

        /// <summary>
        /// This property is used to set/get the count of total table cells.
        /// </summary>
        /// <returns>A string value.</returns>
        public string NumberOfCells
        {
            get { return numberOfCells; }
            set { numberOfCells = value; }
        }

        /// <summary>
        /// This property is used to set/get the cell width.
        /// </summary>
        /// <returns>A string value.</returns>
        public string CellWidth
        {
            get { return cellWidth; }
            set { cellWidth = value; }
        }

        /// <summary>
        /// This property is used to set/get the table border.
        /// </summary>
        /// <returns>A string value.</returns>
        public string TableBorder
        {
            get { return tableBorder; }
            set { tableBorder = value; }
        }

        /// <summary>
        /// This property is used to set/get the table position.
        /// </summary>
        /// <returns>A string value.</returns>
        public string TablePosition
        {
            get { return tablePosition; }
            set { tablePosition = value; }
        }
    }

    /// <summary>
    /// It represents a table cell.
    /// </summary>
    public class TableCell
    {
        /// <value>
        /// An object of the Parent TableCell class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableCell tableCell;

        /// <value>
        /// A variable of type string to store the inner text of cell.
        /// </value>
        protected internal string text;

        /// <value>
        /// A variable of type string to store the width of cell.
        /// </value>
        protected internal string cellWidth;

        /// <summary>
        /// Instantiate an object of TableCell class.
        /// </summary>
        public TableCell()
        {
            this.tableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
        }

        /// <summary>
        /// Invoke this method to add a paragraph into a table cell.
        /// </summary>
        /// <param name="para">An object of the Paragraph class.</param>
        public void Append(FileFormat.Words.Paragraph para)
        {
            this.tableCell.Append(para.wordDocumentParagraph);
        }

        /// <summary>
        /// Call this method to attach properties to the table cell.
        /// </summary>
        /// <param name="tc">An object of the TableCellProperties class.</param>
        public void Append(TableCellProperties tc)
        {
            this.tableCell.Append(tc.cellProperties);
        }

        /// <summary>
        /// This property is used to get/set the width of the cell.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public string CellWidth
        {
            get
            {
                if (this.cellWidth != null)
                    return this.cellWidth;
                if ((this.tableCell.TableCellProperties == null))
                    return "0";
                return this.tableCell.TableCellProperties.TableCellWidth.Width;
            }
            set
            {
                this.cellWidth = value;
            }
        }

        /// <summary>
        /// This property is used to get/set the text of the cell.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public string Text
        {
            get { return this.text; }
            set { this.text = value; }
        }

    }

    /// <summary>
    /// This class contains methods and propertise to merge table cells horizontally.
    /// </summary>
    public class HorizontalMerge
    {
        /// <value>
        /// An object of the Parent HorizontalMerge class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.HorizontalMerge horizontalMerge;

        /// <summary>
        /// Create an object of the HorizontalMerge class.
        /// </summary>
        public HorizontalMerge()
        {
            this.horizontalMerge = new DocumentFormat.OpenXml.Wordprocessing.HorizontalMerge();
        }

        /// <summary>
        /// This property is used to specify that the element shall start a new horizontally merged region in the table.
        /// </summary>
        /// <returns>Returns a boolean value.</returns>
        public bool MergeRestart
        {
            get { return (this.horizontalMerge.Val != null); }
            set
            {
                this.horizontalMerge.Val = DocumentFormat.OpenXml.Wordprocessing.MergedCellValues.Restart;
            }
        }

        /// <summary>
        /// This property is used to specify that the element shall end a horizontally merged region in the table.
        /// </summary>
        /// <returns>Returns a boolean value.</returns>
        public bool MergeContinue
        {
            get { return (this.horizontalMerge.Val != null); }
            set
            {
                this.horizontalMerge.Val = DocumentFormat.OpenXml.Wordprocessing.MergedCellValues.Continue;
            }
        }
    }

    /// <summary>
    /// This class contains methods and propertise to merge table cells vertically.
    /// </summary>
    public class VerticalMerge
    {
        /// <value>
        /// An object of the Parent VerticalMerge class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.VerticalMerge verticalMerge;

        /// <summary>
        /// Create an object of the VerticalMerge class.
        /// </summary>
        public VerticalMerge()
        {
            this.verticalMerge = new DocumentFormat.OpenXml.Wordprocessing.VerticalMerge();
        }

        /// <summary>
        /// This property is used to specify that the element shall start a new vertically merged region in the table.
        /// </summary>
        /// <returns>Returns a boolean value.</returns>
        public bool MergeRestart
        {
            get { return (this.verticalMerge.Val != null); }
            set
            {
                this.verticalMerge.Val = DocumentFormat.OpenXml.Wordprocessing.MergedCellValues.Restart;
            }
        }

        /// <summary>
        /// This property is used to specify that the element shall end a vertically merged region in the table.
        /// </summary>
        /// <returns>Returns a boolean value.</returns>
        public bool MergeContinue
        {
            get { return (this.verticalMerge.Val != null); }
            set
            {
                this.verticalMerge.Val = DocumentFormat.OpenXml.Wordprocessing.MergedCellValues.Continue;
            }
        }

    }

    /// <summary>
    /// This class contains methods to set various table properties.
    /// </summary>
    public class TableCellProperties
    {
        /// <value>
        /// An object of the Parent TableCellProperties class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableCellProperties cellProperties;

        /// <summary>
        /// Create an object of the TableCellProperties class.
        /// </summary>
        public TableCellProperties()
        {
            this.cellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties();
        }

        /// <summary>
        /// Call this method to attach the cell width to TableCellProperties's object.
        /// </summary>
        /// <param name="cellWidth">An instance of the TableJustification class.</param>
        public void Append(TableCellWidth cellWidth)
        {
            this.cellProperties.Append(cellWidth.cellWidth);
        }

        /// <summary>
        /// Call this method to attach an object of the HorizontalMerge class to TableCellProperties's object.
        /// </summary>
        /// <param name="horizontalMerge">An instance of the HorizontalMerge class.</param>
        public void Append(HorizontalMerge horizontalMerge)
        {
            this.cellProperties.Append(horizontalMerge.horizontalMerge);
        }

        /// <summary>
        /// Call this method to attach an object of the VerticalMerge class to TableCellProperties's object.
        /// </summary>
        /// <param name="verticalMerge">An instance of the VerticalMerge class.</param>
        public void Append(VerticalMerge verticalMerge)
        {
            this.cellProperties.Append(verticalMerge.verticalMerge);
        }
    }

    /// <summary>
    /// Represents a table row.
    /// </summary>
    public class TableRow
    {
        /// <value>
        /// An object of the Parent TableRow class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableRow tableRow;

        /// <value>
        /// A variable of type string to store the number of cells. 
        /// </value>
        protected internal string numberOfCell;

        /// <summary>
        /// Instantiate an object of the TableRow class.
        /// </summary>
        public TableRow()
        {
            this.tableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
        }

        /// <summary>
        /// Invoke this method to add cells into a row.
        /// </summary>
        /// <param name="tc">An object of the TableCell class.</param>
        public void Append(TableCell tc)
        {
            this.tableRow.Append(tc.tableCell);
        }

        /// <summary>
        /// This property is used to get/set the number of cells.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public string NumberOfCell
        {
            get { return this.numberOfCell; }
            set { this.numberOfCell = value; }
        }
    }

    /// <summary>
    /// Represents table properties.
    /// </summary>
    public class TableProperties
    {
        /// <value>
        /// An object of the Parent TableProperties class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableProperties tableProperties;

        /// <summary>
        /// Create an instance of the TableProperties class.
        /// </summary>
        public TableProperties()
        {
            this.tableProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties();
        }

        /// <summary>
        /// This method is used to attach the object of TableBorders class to the TableProperties's object.
        /// </summary>
        /// <param name="tableBorders">An instance of the TableBorders class.</param>
        public void Append(TableBorders tableBorders)
        {
            this.tableProperties.Append(tableBorders.tableBorders);
        }

        /// <summary>
        /// Call this method to attach the position of table to the TableProperties's object.
        /// </summary>
        /// <param name="tableJustification">An instance of the TableJustification class.</param>
        public void Append(TableJustification tableJustification)
        {
            this.tableProperties.Append(tableJustification.justification);
        }

    }

    /// <summary>
    /// Represents the border of the table.
    /// </summary>
    public class TableBorders
    {
        /// <value>
        /// An object of the Parent TableBorders class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableBorders tableBorders;

        /// <summary>
        /// Create an object of the TableBorders class.
        /// </summary>
        public TableBorders()
        {
            this.tableBorders = new DocumentFormat.OpenXml.Wordprocessing.TableBorders();
        }

        /// <summary>
        /// This method is used to set the border on the upper side of the table.
        /// </summary>
        /// <param name="topBorder">An instance of the TopBorder class.</param>
        public void AppendTopBorder(TopBorder topBorder) {
            this.tableBorders.Append(topBorder.topBorder);
        }

        /// <summary>
        /// This method is used to set the border on the lower side of the table.
        /// </summary>
        /// <param name="bottomBorder">An instance of the BottomBorder class.</param>
        public void AppendBottomBorder(BottomBorder bottomBorder)
        {
            this.tableBorders.Append(bottomBorder.bottomBorder);
        }

        /// <summary>
        /// This method is used to set the border on the right side of the table.
        /// </summary>
        /// <param name="rightBorder">An instance of the RightBorder class.</param>
        public void AppendRightBorder(RightBorder rightBorder)
        {
            this.tableBorders.Append(rightBorder.rightBorder);
        }

        /// <summary>
        /// This method is used to set the border on the left side of the table.
        /// </summary>
        /// <param name="leftBorder">An instance of the LeftBorder class.</param>
        public void AppendLeftBorder(LeftBorder leftBorder)
        {
            this.tableBorders.Append(leftBorder.leftBorder);
        }

        /// <summary>
        /// This method is used to set the vertical border of the table columns.
        /// </summary>
        /// <param name="insideVerticalBorder">An instance of the InsideVerticalBorder class.</param>
        public void AppendInsideVerticalBorder(InsideVerticalBorder insideVerticalBorder)
        {
            this.tableBorders.Append(insideVerticalBorder.insideVerticalBorder);
        }
        /// <summary>
        /// This method is used to set the horizontal border of the table columns.
        /// </summary>
        /// <param name="insideHorizontalBorder">An instance of the InsideHorizontalBorder class.</param>
        public void AppendInsideHorizontalBorder(InsideHorizontalBorder insideHorizontalBorder)
        {
            this.tableBorders.Append(insideHorizontalBorder.insideHorizontalBorder);
        }
    }

    /// <summary>
    /// Represents the border of the table's upper side.
    /// </summary>
    public class TopBorder
    {
        /// <value>
        /// An object of the Parent TopBorder class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.TopBorder topBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        /// <summary>
        /// Create an object of the TopBorder class.
        /// </summary>
        public TopBorder()
        {
            this.topBorder = new DocumentFormat.OpenXml.Wordprocessing.TopBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }

        /// <summary>
        /// This method is used to set the table's border style to dashed.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dashed_border(UInt32 size)
        {
            this.topBorder.Val = dashed;
            this.topBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to dotted.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dotted_border(UInt32 size)
        {
            this.topBorder.Val = dotted;
            this.topBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to black square.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void basicBlackSquares_border(UInt32 size)
        {
            this.topBorder.Val = black;
            this.topBorder.Size = size;
        }
    }

    /// <summary>
    /// Represents the border of the table's lower side.
    /// </summary>
    public class BottomBorder
    {
        /// <value>
        /// An object of the Parent BottomBorder class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.BottomBorder bottomBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        /// <summary>
        /// Create an object of the BottomBorder class.
        /// </summary>
        public BottomBorder()
        {
            this.bottomBorder = new DocumentFormat.OpenXml.Wordprocessing.BottomBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }

        /// <summary>
        /// This method is used to set the table's border style to dashed.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dashed_border(UInt32 size)
        {
            this.bottomBorder.Val = dashed;
            this.bottomBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to dotted.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dotted_border(UInt32 size)
        {
            this.bottomBorder.Val = dotted;
            this.bottomBorder.Size = size;
        }
        /// <summary>
        /// This method is used to set the table's border style to black square.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void basicBlackSquares_border(UInt32 size)
        {
            this.bottomBorder.Val = black;
            this.bottomBorder.Size = size;
        }
    }

    /// <summary>
    /// Represents the border of the table's right side.
    /// </summary>
    public class RightBorder
    {
        /// <value>
        /// An object of the Parent RightBorder class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.RightBorder rightBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        /// <summary>
        /// Create an object of the RightBorder class.
        /// </summary>
        public RightBorder()
        {
            this.rightBorder = new DocumentFormat.OpenXml.Wordprocessing.RightBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }

        /// <summary>
        /// This method is used to set the table's border style to dashed.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dashed_border(UInt32 size)
        {
            this.rightBorder.Val = dashed;
            this.rightBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to dotted.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dotted_border(UInt32 size)
        {
            this.rightBorder.Val = dotted;
            this.rightBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to black square.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void basicBlackSquares_border(UInt32 size)
        {
            this.rightBorder.Val = black;
            this.rightBorder.Size = size;
        }
    }

    /// <summary>
    /// Represents the border of the table's left side.
    /// </summary>
    public class LeftBorder
    {
        /// <value>
        /// An object of the Parent LeftBorder class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.LeftBorder leftBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        /// <summary>
        /// Create an object of the LeftBorder class.
        /// </summary>
        public LeftBorder()
        {
            this.leftBorder = new DocumentFormat.OpenXml.Wordprocessing.LeftBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }

        /// <summary>
        /// This method is used to set the table's border style to dashed.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dashed_border(UInt32 size)
        {
            this.leftBorder.Val = dashed;
            this.leftBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to dotted.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dotted_border(UInt32 size)
        {
            this.leftBorder.Val = dotted;
            this.leftBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to black square.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void basicBlackSquares_border(UInt32 size)
        {
            this.leftBorder.Val = black;
            this.leftBorder.Size = size;
        }
    }

    /// <summary>
    /// Represents the inner vertical border of the table.
    /// </summary>
    public class InsideVerticalBorder
    {
        /// <value>
        /// An object of the Parent InsideVerticalBorder class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder insideVerticalBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        /// <summary>
        /// Create an object of the InsideVerticalBorder class.
        /// </summary>
        public InsideVerticalBorder()
        {
            this.insideVerticalBorder = new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }

        /// <summary>
        /// This method is used to set the table's border style to dashed.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dashed_border(UInt32 size)
        {
            this.insideVerticalBorder.Val = dashed;
            this.insideVerticalBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to dotted.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dotted_border(UInt32 size)
        {
            this.insideVerticalBorder.Val = dotted;
            this.insideVerticalBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to black square.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void basicBlackSquares_border(UInt32 size)
        {
            this.insideVerticalBorder.Val = black;
            this.insideVerticalBorder.Size = size;
        }
    }

    /// <summary>
    /// Represents the inner horizontal border of the table.
    /// </summary>
    public class InsideHorizontalBorder
    {
        /// <value>
        /// An object of the Parent InsideHorizontalBorder class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder insideHorizontalBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        /// <summary>
        /// Create an object of the InsideHorizontalBorder class.
        /// </summary>
        public InsideHorizontalBorder()
        {
            this.insideHorizontalBorder = new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }

        /// <summary>
        /// This method is used to set the table's border style to dashed.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dashed_border(UInt32 size)
        {
            this.insideHorizontalBorder.Val = dashed;
            this.insideHorizontalBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to dotted.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void dotted_border(UInt32 size)
        {
            this.insideHorizontalBorder.Val = dotted;
            this.insideHorizontalBorder.Size = size;
        }

        /// <summary>
        /// This method is used to set the table's border style to black square.
        /// </summary>
        /// <param name="size">An integer value that represents the border thickness.</param>
        public void basicBlackSquares_border(UInt32 size)
        {
            this.insideHorizontalBorder.Val = black;
            this.insideHorizontalBorder.Size = size;
        }
    }

    /// <summary>
    /// Represents the width of table cells.
    /// </summary>
    public class TableCellWidth
    {
        /// <value>
        /// An object of the Parent TableCellWidth class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableCellWidth cellWidth;

        /// <summary>
        /// Invoke this method to set the cell width.
        /// </summary>
        /// <param name="width">An instance of the TableJustification class.</param>
        public TableCellWidth(string width)
        {
            this.cellWidth = new DocumentFormat.OpenXml.Wordprocessing.TableCellWidth() { Width = width };
        }

        /// <summary>
        /// This property is used to get the width of the cell.
        /// </summary>
        /// <returns>Returns a string value.</returns>
        public string CellWidth { get { return this.cellWidth.Width; } }

    }

    /// <summary>
    /// The class helps set the position of the table.
    /// </summary>
    public class TableJustification
    {
        /// <value>
        /// An object of the Parent TableJustification class.
        /// </value>
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableJustification justification;
        private DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues center;
        private DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues right;
        private DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues left;

        /// <summary>
        /// Initialize an object of the TableJustification class.
        /// </summary>
        public TableJustification()
        {
            this.justification = new DocumentFormat.OpenXml.Wordprocessing.TableJustification();
            center = DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues.Center;
            right = DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues.Right;
            left = DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues.Left;
        }

        /// <summary>
        /// Invoke this method to align the table to the center position in the document.
        /// </summary>
        /// <param name="tableRowAlignmentValues">An instance of the TableRowAlignmentValues class.</param>
        public void AlignCneter() {
            this.justification.Val = center;
        }

        /// <summary>
        /// Invoke this method to align the table to the left position in the document.
        /// </summary>
        /// <param name="tableRowAlignmentValues">An instance of the TableRowAlignmentValues class.</param>
        public void AlignLeft()
        {
            this.justification.Val = left;
        }

        /// <summary>
        /// Invoke this method to align the table to the right position in the document.
        /// </summary>
        /// <param name="tableRowAlignmentValues">An instance of the TableRowAlignmentValues class.</param>
        public void AlignRight()
        {
            this.justification.Val = right;
        }

    }

}

