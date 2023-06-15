using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace FileFormat.Words.Table
{
    public class Table
    {
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

        public Table()
        {
            this.table = new DocumentFormat.OpenXml.Wordprocessing.Table();
        }
        public void TableHeaders(string val)
        {
            tableHeaders.Add(val);
        }
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
        
        public List<string> ExistingTableHeaders
        {
            get { return existingTableHeaders; }
            set
            {
                existingTableHeaders.Add(value.ToString());
            }
        }

        public void AppendChild(TableProperties tp)
        {
            this.table.AppendChild(tp.tableProperties);
        }
        public int getTableRows
        {
            get
            {
                return this.table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ToList().Count;
            }
        }
        public int getTableCells
        {
            get
            {
                return this.table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList().Count;
            }
        }
        public int getTableColumns
        {
            get
            {

                return this.table.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList().Count;
            }
        }
        public string NumberOfTables
        {
            get { return numberOfTables; }
            set { numberOfTables = value; }
        }
        public string NumberOfRows
        {
            get { return numberOfRows; }
            set { numberOfRows = value; }
        }
        public string NumberOfColumns
        {
            get { return numberOfColumns; }
            set { numberOfColumns = value; }
        }
        public string NumberOfCells
        {
            get { return numberOfCells; }
            set { numberOfCells = value; }
        }
        public string CellWidth
        {
            get { return cellWidth; }
            set { cellWidth = value; }
        }
        public string TableBorder
        {
            get { return tableBorder; }
            set { tableBorder = value; }
        }
        public string TablePosition
        {
            get { return tablePosition; }
            set { tablePosition = value; }
        }
    }

    public class TableCell
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableCell tableCell;
        protected internal string text;
        protected internal string cellWidth;

        public TableCell()
        {
            this.tableCell = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
        }
        public void Append(FileFormat.Words.Paragraph para)
        {
            this.tableCell.Append(para.wordDocumentParagraph);
        }
        public void Append(TableCellProperties tc)
        {
            this.tableCell.Append(tc.cellProperties);
        }
        public void setHeader(string header)
        {

        }
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
        public string Text
        {
            get { return this.text; }
            set { this.text = value; }
        }

    }

    public class TableCellProperties
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableCellProperties cellProperties;
        public TableCellProperties()
        {
            this.cellProperties = new DocumentFormat.OpenXml.Wordprocessing.TableCellProperties();
        }
        public void Append(TableCellWidth cellWidth)
        {
            this.cellProperties.Append(cellWidth.cellWidth);
        }
    }
    public class TableRow
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableRow tableRow;
        protected internal string numberOfCell;

        public TableRow()
        {
            this.tableRow = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
        }
        public void Append(TableCell tc)
        {
            this.tableRow.Append(tc.tableCell);
        }

        public string NumberOfCell
        {
            get { return this.numberOfCell; }
            set { this.numberOfCell = value; }
        }
    }

    public class TableProperties
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableProperties tableProperties;

        public TableProperties()
        {
            this.tableProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties();
        }
        public void Append(TableBorders tableBorders)
        {
            this.tableProperties.Append(tableBorders.tableBorders);
        }
        public void Append(TableJustification tableJustification)
        {
            this.tableProperties.Append(tableJustification.justification);
        }

    }
    public class TableBorders
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableBorders tableBorders;

        public TableBorders()
        {
            this.tableBorders = new DocumentFormat.OpenXml.Wordprocessing.TableBorders();
        }

        public void AppendTopBorder(TopBorder topBorder) {
            this.tableBorders.Append(topBorder.topBorder);
        }
        public void AppendBottomBorder(BottomBorder bottomBorder)
        {
            this.tableBorders.Append(bottomBorder.bottomBorder);
        }
        public void AppendRightBorder(RightBorder rightBorder)
        {
            this.tableBorders.Append(rightBorder.rightBorder);
        }
        public void AppendLeftBorder(LeftBorder leftBorder)
        {
            this.tableBorders.Append(leftBorder.leftBorder);
        }
        public void AppendInsideVerticalBorder(InsideVerticalBorder insideVerticalBorder)
        {
            this.tableBorders.Append(insideVerticalBorder.insideVerticalBorder);
        }
        public void AppendInsideHorizontalBorder(InsideHorizontalBorder insideHorizontalBorder)
        {
            this.tableBorders.Append(insideHorizontalBorder.insideHorizontalBorder);
        }
    }
    public class TopBorder
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TopBorder topBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        public TopBorder()
        {
            this.topBorder = new DocumentFormat.OpenXml.Wordprocessing.TopBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }
        public void dashed_border(UInt32 size)
        {
            this.topBorder.Val = dashed;
            this.topBorder.Size = size;
        }
        public void dotted_border(UInt32 size)
        {
            this.topBorder.Val = dotted;
            this.topBorder.Size = size;
        }
        public void basicBlackSquares_border(UInt32 size)
        {
            this.topBorder.Val = black;
            this.topBorder.Size = size;
        }
    }
    public class BottomBorder
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.BottomBorder bottomBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        public BottomBorder()
        {
            this.bottomBorder = new DocumentFormat.OpenXml.Wordprocessing.BottomBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }
        public void dashed_border(UInt32 size)
        {
            this.bottomBorder.Val = dashed;
            this.bottomBorder.Size = size;
        }
        public void dotted_border(UInt32 size)
        {
            this.bottomBorder.Val = dotted;
            this.bottomBorder.Size = size;
        }
        public void basicBlackSquares_border(UInt32 size)
        {
            this.bottomBorder.Val = black;
            this.bottomBorder.Size = size;
        }
    }
    public class RightBorder
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.RightBorder rightBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        public RightBorder()
        {
            this.rightBorder = new DocumentFormat.OpenXml.Wordprocessing.RightBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }
        public void dashed_border(UInt32 size)
        {
            this.rightBorder.Val = dashed;
            this.rightBorder.Size = size;
        }
        public void dotted_border(UInt32 size)
        {
            this.rightBorder.Val = dotted;
            this.rightBorder.Size = size;
        }
        public void basicBlackSquares_border(UInt32 size)
        {
            this.rightBorder.Val = black;
            this.rightBorder.Size = size;
        }
    }
    public class LeftBorder
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.LeftBorder leftBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        public LeftBorder()
        {
            this.leftBorder = new DocumentFormat.OpenXml.Wordprocessing.LeftBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }
        public void dashed_border(UInt32 size)
        {
            this.leftBorder.Val = dashed;
            this.leftBorder.Size = size;
        }
        public void dotted_border(UInt32 size)
        {
            this.leftBorder.Val = dotted;
            this.leftBorder.Size = size;
        }
        public void basicBlackSquares_border(UInt32 size)
        {
            this.leftBorder.Val = black;
            this.leftBorder.Size = size;
        }
    }
    public class InsideVerticalBorder
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder insideVerticalBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        public InsideVerticalBorder()
        {
            this.insideVerticalBorder = new DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }
        public void dashed_border(UInt32 size)
        {
            this.insideVerticalBorder.Val = dashed;
            this.insideVerticalBorder.Size = size;
        }
        public void dotted_border(UInt32 size)
        {
            this.insideVerticalBorder.Val = dotted;
            this.insideVerticalBorder.Size = size;
        }
        public void basicBlackSquares_border(UInt32 size)
        {
            this.insideVerticalBorder.Val = black;
            this.insideVerticalBorder.Size = size;
        }
    }
    public class InsideHorizontalBorder
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder insideHorizontalBorder;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dashed;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues dotted;
        private DocumentFormat.OpenXml.Wordprocessing.BorderValues black;

        public InsideHorizontalBorder()
        {
            this.insideHorizontalBorder = new DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder();
            dashed = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed;
            dotted = DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted;
            black = DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares;
        }
        public void dashed_border(UInt32 size)
        {
            this.insideHorizontalBorder.Val = dashed;
            this.insideHorizontalBorder.Size = size;
        }
        public void dotted_border(UInt32 size)
        {
            this.insideHorizontalBorder.Val = dotted;
            this.insideHorizontalBorder.Size = size;
        }
        public void basicBlackSquares_border(UInt32 size)
        {
            this.insideHorizontalBorder.Val = black;
            this.insideHorizontalBorder.Size = size;
        }
    }
    public class BorderValues
    {
        public DocumentFormat.OpenXml.Wordprocessing.BorderValues borderValue_Dashed
        {
            get { return DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dashed; }
        }
        public DocumentFormat.OpenXml.Wordprocessing.BorderValues borderValue_Dotted
        {
            get { return DocumentFormat.OpenXml.Wordprocessing.BorderValues.Dotted; }
        }

        public DocumentFormat.OpenXml.Wordprocessing.BorderValues borderValue_BasicBlackSquares
        {
            get { return DocumentFormat.OpenXml.Wordprocessing.BorderValues.BasicBlackSquares; }
        }
    }

    public class TableCellWidth
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableCellWidth cellWidth;

        public TableCellWidth(string width)
        {
            this.cellWidth = new DocumentFormat.OpenXml.Wordprocessing.TableCellWidth() { Width = width };
        }
        public string CellWidth { get { return this.cellWidth.Width; } }

    }
    public class TableJustification
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableJustification justification;
        private DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues center;
        private DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues right;
        private DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues left;

        public TableJustification()
        {
            this.justification = new DocumentFormat.OpenXml.Wordprocessing.TableJustification();
            center = DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues.Center;
            right = DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues.Right;
            left = DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues.Left;
        }
        public void AlignCneter() {
            this.justification.Val = center;
        }
        public void AlignLeft()
        {
            this.justification.Val = left;
        }
        public void AlignRight()
        {
            this.justification.Val = right;
        }

    }

}

