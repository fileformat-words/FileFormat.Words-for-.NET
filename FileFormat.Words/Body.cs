using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using FileFormat.Words.Table;


namespace FileFormat.Words
{
    public class Body : Document
    {
        protected internal DocumentFormat.OpenXml.Wordprocessing.Body wordDocumentBody;

        public Body(Document doc)
        {
            this.wordDocumentBody = doc.wordDocument.MainDocumentPart.Document.Body;
        }

        public void AppendChild(Paragraph para)
        {
            this.wordDocumentBody.AppendChild(para.wordDocumentParagraph);
        }
        public void AppendChild(Table.Table tab)
        {
            this.wordDocumentBody.Append(tab.table);
        }
        public void LineBreak()
        {
            this.wordDocumentBody.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new Text(" "))));
        }
        public List<Paragraph> GetParagraphs()
        {
            var paragraphs = this.wordDocumentBody.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>();
            List<Paragraph> lst = new List<Paragraph>();
            Paragraph para1;


            foreach (DocumentFormat.OpenXml.Wordprocessing.Paragraph para in paragraphs.ToArray())
            {
                if (para.InnerText != " ")
                {
                    para1 = new Paragraph();
                    para1.wordDocumentParagraph.ParagraphProperties = new ParagraphProperties();
                    para1.Text = para.InnerText;
                    para1.Indent = (para.ParagraphProperties != null && para.ParagraphProperties.Indentation != null) ? para.ParagraphProperties.Indentation.Start.Value : null;
                    para1.LinesSpacing = (para.ParagraphProperties != null && para.ParagraphProperties.SpacingBetweenLines != null) ? para.ParagraphProperties.SpacingBetweenLines.Line.Value : null;

                    lst.Add(para1);
                }

            }
            return lst;
        }
        public List<Table.Table> getDocumentTables
        {

            get
            {
                List<Table.Table> ls = new List<Table.Table>();
                Table.Table table = new Table.Table();
;
                foreach (DocumentFormat.OpenXml.Wordprocessing.Table tbl in this.wordDocumentBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList())
                {

                    Table.Table tab = new Table.Table();
                    for (int i = 0; i < tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>().First().Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count(); i++)
                    {
                        tab.ExistingTableHeaders.Add(tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>().First().Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList()[i].InnerText);
                    }
                    tab.NumberOfRows = tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Count().ToString();
                    tab.NumberOfColumns = tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableRow>().First().Count().ToString();
                    tab.NumberOfCells = tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ToList().Count().ToString();
                    if (tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableJustification>().ToList().Count() > 0)
                        tab.TablePosition = tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableJustification>().ToList()[0].Val;
                    else tab.TablePosition = "null";
                    if (tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().First().TableCellProperties != null)
                        tab.CellWidth = int.Parse(tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCell>().First().TableCellProperties.TableCellWidth.Width).ToString();
                    else tab.CellWidth = "null";
                    if ((this.wordDocumentBody).Descendants<DocumentFormat.OpenXml.Wordprocessing.TableBorders>().Count() != 0)
                        tab.TableBorder = tbl.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableBorders>().FirstOrDefault().Descendants().First().GetAttributes()[0].Value;
                    else tab.TableBorder = "null";
                    ls.Add(tab);
                }

                return ls;
            }
        }
        public List<Table.TableCell> FindTableCell(int tableIndex, int tableRow, int tableCell)
        {

            List<Table.TableCell> ls = new List<Table.TableCell>();
            Table.Table tab = new Table.Table();
            Table.TableCell tableCell1 = new Table.TableCell();
            DocumentFormat.OpenXml.Wordprocessing.Table table = this.wordDocumentBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList()[tableIndex];
            DocumentFormat.OpenXml.Wordprocessing.TableRow row = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ElementAt(tableRow);
            if (tableIndex >= this.wordDocumentBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList().Count())
            {

                tableCell1.CellWidth = null;
                tableCell1.Text = null;
                ls.Add(tableCell1);
                return ls;
            }
            if (tableRow >= table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Count())
            {

                tableCell1.CellWidth = null;
                tableCell1.Text = null;
                ls.Add(tableCell1);
                return ls;
            }
            if (tableCell >= row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count())
            {

                tableCell1.CellWidth = null;
                tableCell1.Text = null;
                ls.Add(tableCell1);
                return ls;
            }

            DocumentFormat.OpenXml.Wordprocessing.TableCell cell = row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAt(tableCell);

            tableCell1.CellWidth = cell.TableCellProperties.TableCellWidth.Width.ToString();
            tableCell1.Text = cell.InnerText;
            ls.Add(tableCell1);
            return ls;
        }
        public List<Table.TableRow> FindTableRow(int tableindex, int tableRowIndex)
        {
            Table.Table tab = new Table.Table();
            List<Table.TableRow> ls = new List<Table.TableRow>();
            Table.TableRow tableRow1 = new Table.TableRow();
            if (tableindex >= this.wordDocumentBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList().Count())
            {

                tableRow1.NumberOfCell = null;
                ls.Add(tableRow1);
                return ls;
            }

            DocumentFormat.OpenXml.Wordprocessing.Table table = this.wordDocumentBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList()[tableindex];
            if (tableRowIndex >= table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().Count())
            {
                tableRow1.NumberOfCell = null;
                ls.Add(tableRow1);
                return ls;
            }
            DocumentFormat.OpenXml.Wordprocessing.TableRow row = table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>().ElementAt(tableRowIndex);
            tableRow1.NumberOfCell = row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().Count().ToString();

            ls.Add(tableRow1);
            return ls;
        }
        public int FindTableByText(string text)
        {
            List<DocumentFormat.OpenXml.Wordprocessing.Table> table = wordDocumentBody.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
            List<DocumentFormat.OpenXml.Wordprocessing.Table> ls = new List<DocumentFormat.OpenXml.Wordprocessing.Table>();
            foreach (DocumentFormat.OpenXml.Wordprocessing.Table tab in table)
            {
                if (tab.InnerText.Contains(text))
                    ls.Add(tab);
            }
            return ls.Count();
        }
    }
}

