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
                string rootDir =  "/Users/Mustafa/Projects/FileFormat.Words/TestDocs";
    
                if (!Directory.Exists(rootDir))
                {
                    Directory.CreateDirectory(rootDir);
                }
            

                using (Document doc = new Document())
                {

                    Body body = new Body(doc);

                    // ********** this code block creates a paragraph into a word document  ********** // 
                    Paragraph para1 = new Paragraph();
                    para1.Style = "Heading2";
                    para1.Text = "The Compare operation is particularly useful. For example, if I took my test Word document and saved it before setting the paragraph setting to and then set the setting to and saved another copy, then I could compare the two documents and that one change would be highlighted.";
                    para1.Indent = "300";
                    para1.LeftIndent = "250";
                    para1.RihgtIndent = "350";
                    para1.FirstLineIndent = "330";
                    para1.Align = "Left";
                    para1.LinesSpacing = "552";

                    body.AppendChild(para1);

                    Run run1 = new Run();
                    run1.Bold = true;
                    run1.Italic = true;
                    run1.FontFamily = "Algerian";
                    run1.FontSize = 40;
                    run1.Underline = true;
                    run1.Color = "FF0000";
                    run1.Text = "Text for the run";


                    para1.AppendChild(run1);

                    IEnumerable<Run> runs = para1.GetRuns();

                    //Loop through each run and print its text
                    foreach (Run runner in runs)
                    {
                        Console.WriteLine("Runner Text = " + runner.Text);
                    }

                    // insert a new line into a document 
                    body.LineBreak();


                    // ********** this code block creates a table into a word document  ********** // 
                    Table table = new Table();

                    BorderValues borderValues = new BorderValues();
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
                    tblCellProps2.Append(new TableCellWidth("1400"));
                    tableCell2.Append(tblCellProps2);

                    TableCell tableCell3 = new TableCell();
                    Paragraph para3 = new Paragraph();
                    Run run3 = new Run();
                    table.TableHeaders("Age");
                    run3.Text = "30";
                    para3.AppendChild(run3);
                    tableCell3.Append(para3);

                    TableCellProperties tblCellProps3 = new TableCellProperties();
                    tblCellProps3.Append(new TableCellWidth("1400"));
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



                    // ********** this code block adds images to an existing Word document ********** // 

                    var imagePath = rootDir + "/img1.png";
                    var image = new Image(doc, imagePath, 100, 100);
                    var paragraph = new Paragraph();
                    var imageRun = new Run();
                    imageRun.AppendChild(image);
                    paragraph.AppendChild(imageRun);
                    body.AppendChild(paragraph);

                    // ********** end of image creation code block  ********** //

                    doc.Save(rootDir + "/Docs.docx");

                    Console.WriteLine(" Document created successfully ");
                }

                // ********** this code block reads tables from an existing Word document ********** // 
                using (Document doc1 = new Document(rootDir + "/Docs2.docx"))
                {
                    Body body1 = new Body(doc1);
                    Console.WriteLine("Total Number of Tables " + body1.getDocumentTables.Count());

                    // find table by text
                    int tableCount = body1.FindTableByText("Name");
                    Console.WriteLine("number of tables with this text = " + tableCount);

                    foreach (Table props in body1.getDocumentTables)
                    {
                        foreach (string tableHeader in props.ExistingTableHeaders)
                        {
                            Console.WriteLine(tableHeader);
                        }
                        Console.WriteLine(props.NumberOfRows);
                        Console.WriteLine(props.NumberOfColumns);
                        Console.WriteLine(props.NumberOfCells);
                        Console.WriteLine(props.CellWidth);
                        Console.WriteLine(props.TableBorder);
                        Console.WriteLine(props.TablePosition);
                        Console.WriteLine(" ");

                    }
                    foreach (FileFormat.Words.Table.TableRow row in body1.FindTableRow(0, 1))
                    {
                        Console.WriteLine(row.NumberOfCell);
                    }
                    foreach (FileFormat.Words.Table.TableCell cell in body1.FindTableCell(0, 1, 1))
                    {
                        Console.WriteLine(cell.Text);
                        Console.WriteLine(cell.CellWidth);
                    }
                    Table tble = new Table();
                    Console.WriteLine(tble.ChangeTextInCell(rootDir + "/Docs2.docx", 0, 1, 2, "changed"));

                    // ********** this code block reads paragraphs from an existing Word document ********** // 
                    List<Paragraph> paras = body1.GetParagraphs();

                    Console.WriteLine("The number of Paragraphs " + paras.Count());
                    foreach (Paragraph p in paras)
                    {
                        Console.WriteLine(p.LinesSpacing);
                        Console.WriteLine(p.Indent);
                        Console.WriteLine(p.Text);
                    }

                    // ********** this code block reads images from an existing Word document ********** //

                    List<Stream> imageParts = Image.ExtractImagesFromDocument(doc1);
                    int imageCount = imageParts.Count;
                    Console.WriteLine($"Total number of images: {imageCount}");
                    //// Process the image parts as needed
                    var i = 1;
                    foreach (Stream imagePart in imageParts)
                    {
                        using (FileStream stream = new FileStream(rootDir + "/" + i + ".jpeg", FileMode.Create))
                        {
                            imagePart.CopyTo(stream);
                        }
                        i = i + 1;
                    }
                    //Console.ReadLine();
                }

            }
            catch (Exception e)
            {
                throw e;
            }

        }

    }
}
