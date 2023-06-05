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
        public List<DocumentFormat.OpenXml.Wordprocessing.Table> getDocumentTablesByText(DocumentFormat.OpenXml.Wordprocessing.Body wordDocumentBody, string txt)
        {
            List<DocumentFormat.OpenXml.Wordprocessing.Table> table = wordDocumentBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
            List<DocumentFormat.OpenXml.Wordprocessing.Table> ls = new List<DocumentFormat.OpenXml.Wordprocessing.Table>();
            foreach (DocumentFormat.OpenXml.Wordprocessing.Table tab in table)
            {
                if (tab.InnerText.Contains(txt))
                    ls.Add(tab);
            }
            return ls;
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

        public DocumentFormat.OpenXml.Wordprocessing.BorderValues borderValue_Dashed
        {
            get { return BorderValues.Dashed; }
        }
        public DocumentFormat.OpenXml.Wordprocessing.BorderValues borderValue_Dotted
        {
            get { return BorderValues.Dotted; }
        }

        public DocumentFormat.OpenXml.Wordprocessing.BorderValues borderValue_BasicBlackSquares
        {
            get { return BorderValues.BasicBlackSquares; }
        }
        public DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues AlignCenter
        {
            get { return TableRowAlignmentValues.Center; }
        }
        public DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues AlignLeft
        {
            get { return TableRowAlignmentValues.Left; }
        }
        public DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues AlignRight
        {
            get { return TableRowAlignmentValues.Right; }
        }
        public List<DocumentFormat.OpenXml.Wordprocessing.Table> getDocumentTables(DocumentFormat.OpenXml.Wordprocessing.Body wordDocumentBody)
        {
            List<DocumentFormat.OpenXml.Wordprocessing.Table> table = wordDocumentBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
            return table;
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
        public void Append(DocumentFormat.OpenXml.Wordprocessing.TableRowProperties tc)
        {
            this.tableRow.Append(tc);
        }
        public string NumberOfCell
        {
            get { return this.numberOfCell; }
            set { this.numberOfCell = value; }
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

    public class TableProperties
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.TableProperties tableProperties;

        public TableProperties(UInt32 size, DocumentFormat.OpenXml.Wordprocessing.BorderValues borderValues)
        {
            this.tableProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(new TableBorders(
                new TopBorder()
                {
                    Val =
            new EnumValue<BorderValues>(borderValues),
                    Size = size
                },
                new BottomBorder()
                {
                    Val =
            new EnumValue<BorderValues>(borderValues),
                    Size = size
                },
                new LeftBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(borderValues),
                    Size = size
                },
                new RightBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(borderValues),
                    Size = size
                },

                new InsideVerticalBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(borderValues),
                    Size = size
                },

                new InsideHorizontalBorder()
                {
                    Val =
                    new EnumValue<BorderValues>(borderValues),
                    Size = size
                }
            ));
        }

        public void Append(TableJustification tableJustification)
        {
            this.tableProperties.Append(tableJustification.justification);
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
        public TableJustification(DocumentFormat.OpenXml.Wordprocessing.TableRowAlignmentValues tableRowAlignmentValues)
        {
            this.justification = new DocumentFormat.OpenXml.Wordprocessing.TableJustification() { Val = tableRowAlignmentValues };
        }

    }

}

