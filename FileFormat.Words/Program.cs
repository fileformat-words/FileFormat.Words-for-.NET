using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using FileFormat.Words;
using FileFormat.Words.Table;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // create directory for reading/writing files. 
                string rootDir = "/Users/Mustafa/Desktop";

                if (!Directory.Exists(rootDir))
                {
                    Directory.CreateDirectory(rootDir);
                }


                using (Document doc = new Document())
                {

                    Body body = new Body(doc);

                    // ********** this code block creates a table into a word document  ********** // 
                    Table table = new Table();

                    TopBorder topBorder = new TopBorder();
                    topBorder.basicBlackSquares_border(20);

                    BottomBorder bottomBorder = new BottomBorder();
                    bottomBorder.basicBlackSquares_border(20);

                    RightBorder rightBorder = new RightBorder();
                    rightBorder.basicBlackSquares_border(20);

                    LeftBorder leftBorder = new LeftBorder();
                    leftBorder.basicBlackSquares_border(20);

                    InsideVerticalBorder insideVerticalBorder = new InsideVerticalBorder();
                    insideVerticalBorder.basicBlackSquares_border(20);

                    InsideHorizontalBorder insideHorizontalBorder = new InsideHorizontalBorder();
                    insideHorizontalBorder.basicBlackSquares_border(20);

                    TableBorders tableBorders = new TableBorders();
                    tableBorders.AppendTopBorder(topBorder);
                    tableBorders.AppendBottomBorder(bottomBorder);
                    tableBorders.AppendRightBorder(rightBorder);
                    tableBorders.AppendLeftBorder(leftBorder);
                    tableBorders.AppendInsideVerticalBorder(insideVerticalBorder);
                    tableBorders.AppendInsideHorizontalBorder(insideHorizontalBorder);

                    // specify its border style of the table
                    TableProperties tblProp = new TableProperties();

                    tblProp.Append(tableBorders);

                    TableJustification tableJustification = new TableJustification();
                    tableJustification.AlignLeft();
                    // set the position of the table to Right.
                    tblProp.Append(tableJustification);

                    // create two table rows
                    TableRow tableRow = new TableRow();
                    TableRow tableRow2 = new TableRow();

                    // create table cell
                    TableCell tableCell = new TableCell();
                    Paragraph para = new Paragraph();
                    Run run = new Run();

                    // set the header of the first column
                    table.TableHeaders("Name");
                    run.Text = "Mustafa";
                    para.AppendChild(run);
                    tableCell.Append(para);

                    // create table properties
                    TableCellProperties tblCellProps = new TableCellProperties();

                    // set the width of table cell 
                    tblCellProps.Append(new TableCellWidth("2400"));
                    tableCell.Append(tblCellProps);

                    TableCell tableCell2 = new TableCell();
                    Paragraph para2 = new Paragraph();
                    Run run2 = new Run();

                    // set the header of the second column
                    table.TableHeaders("Nationality");
                    run2.Text = "Pakistani";
                    para2.AppendChild(run2);
                    tableCell2.Append(para2);

                    TableCellProperties tblCellProps2 = new TableCellProperties();

                    VerticalMerge verticalMerge = new VerticalMerge();
                    verticalMerge.MergeRestart = true;
                    tblCellProps.Append(verticalMerge);

                    HorizontalMerge horizontalMerge = new HorizontalMerge();
                    horizontalMerge.MergeRestart = true;
                    tblCellProps2.Append(horizontalMerge);

                    tblCellProps2.Append(new TableCellWidth("1400"));
                    tableCell2.Append(tblCellProps2);

                    TableCell tableCell3 = new TableCell();
                    Paragraph para3 = new Paragraph();
                    Run run3 = new Run();
                    table.TableHeaders("Age");
                    run3.Text = "30";
                    para3.AppendChild(run3);
                    tableCell3.Append(para3);

                    HorizontalMerge horizontalMerge1 = new HorizontalMerge();
                    horizontalMerge1.MergeContinue = true;
                    TableCellProperties tblCellProps3 = new TableCellProperties();
                    tblCellProps3.Append(new TableCellWidth("1400"));
                    tblCellProps3.Append(horizontalMerge1);
                    tableCell3.Append(tblCellProps3);

                    tableRow.Append(tableCell);
                    tableRow.Append(tableCell2);
                    tableRow.Append(tableCell3);

                    // create table cell
                    TableCell _tableCell = new TableCell();
                    Paragraph _para = new Paragraph();
                    Run _run = new Run();

                    _run.Text = "sultan";
                    _para.AppendChild(_run);
                    _tableCell.Append(_para);

                    TableCellProperties tblCellProps1_ = new TableCellProperties();

                    VerticalMerge verticalMerge2 = new VerticalMerge();
                    verticalMerge2.MergeContinue = true;
                    tblCellProps1_.Append(verticalMerge2);

                    tblCellProps1_.Append(new TableCellWidth("2400"));
                    _tableCell.Append(tblCellProps1_);


                    TableCell _tableCell2 = new TableCell();
                    Paragraph _para2 = new Paragraph();
                    Run _run2 = new Run();

                    _run2.Text = "British";
                    _para2.AppendChild(_run2);
                    _tableCell2.Append(_para2);

                    TableCellProperties tblCellProps2_ = new TableCellProperties();
                    tblCellProps2_.Append(new TableCellWidth("1400"));
                    _tableCell2.Append(tblCellProps2_);

                    TableCell _tableCell3 = new TableCell();
                    Paragraph _para3 = new Paragraph();
                    Run _run3 = new Run();

                    _run3.Text = "2";
                    _para3.AppendChild(_run3);
                    _tableCell3.Append(_para3);

                    TableCellProperties tblCellProps3_ = new TableCellProperties();
                    tblCellProps3_.Append(new TableCellWidth("1400"));
                    _tableCell3.Append(tblCellProps3_);

                    tableRow2.Append(_tableCell);
                    tableRow2.Append(_tableCell2);
                    tableRow2.Append(_tableCell3);

                    table.AppendChild(tblProp);
                    table.Append(tableRow);
                    table.Append(tableRow2);

                    body.AppendChild(table);
                    // ********** end of table creation code block  ********** //


                    doc.Save(rootDir + "/Docs.docx");

                    Console.WriteLine(" Document created successfully ");
                }

            }
            catch (Exception e)
            {
                throw e;
            }

        }

    }
}
