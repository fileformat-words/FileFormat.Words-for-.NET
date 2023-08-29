using Microsoft.VisualStudio.TestTools.UnitTesting;
using FileFormat.Words;
using FileFormat.Words.Properties;
using FileFormat.Words.Table;
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using static System.Net.Mime.MediaTypeNames;
using System.Collections.Generic;

/// <summary>
/// Namspace for Testing FileFormat.Words.Document Class
/// </summary>
namespace FileFormat_Tests
{
    /// <summary>
    /// Class for Testing FileFormat.Words.Document Class
    /// </summary>
    [TestClass]
    public class TestDocumentClass
    {

        private static string testDir = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/";
        private static string processedDir = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/ProcessedDocs/";
        private static string testDoc = "UbuntuSoftwareCenter";
        /// <summary>
        /// Test #1 Create empty WordprocessingML document and save to disk
        /// </summary>
        [TestMethod]
        public void TestCreateNSave()
        {
            using (Document doc = new Document())
                doc.Save(processedDir + "Created_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx");
        }
        /// <summary>
        /// Test #2 Create empty WordprocessingML document and save to stream
        /// </summary>
        [TestMethod]
        public void TestCreateNSaveStream()
        {
            using (Document doc = new Document())
            using (FileStream fs = new FileStream(processedDir + "CreatedStream_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx", FileMode.Create))
                doc.Save(fs);
        }
        /// <summary>
        /// Test #3 Open WordprocessingML document from disk and save to disk
        /// </summary>
        [TestMethod]
        public void TestLoadNSave()
        {
            using (Document doc = new Document(testDir + testDoc + ".docx"))
                doc.Save(processedDir + testDoc + "_Saved_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx");
        }
        /// <summary>
        /// Test #4 Open WordprocessingML document from disk and save to stream
        /// </summary>
        [TestMethod]
        public void TestLoadNSaveStream()
        {
            using (Document doc = new Document(testDir + testDoc + ".docx"))
            using (FileStream fs = new FileStream(processedDir + testDoc + "_SavedStream_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx", FileMode.Create))
                doc.Save(fs);
        }
        /// <summary>
        /// Test #5 Open WordprocessingML document from stream and save to disk
        /// </summary>
        [TestMethod]
        public void TestLoadStreamNSave()
        {
            using (FileStream fs = new FileStream(testDir + testDoc + ".docx", FileMode.Open))
            using (Document doc = new Document(fs))
                doc.Save(processedDir + testDoc + "_StreamSaved_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx");
        }
        /// <summary>
        /// Test #6 Get BuiltInDocumentProperties from a WordprocessingML document
        /// </summary>
        [TestMethod]
        public void TestGetBuiltInDocumentProperties()
        {
            BuiltInDocumentProperties prop;
            using (Document doc = new Document(testDir + testDoc + ".docx"))
            {
                prop = doc.BuiltinDocumentProperties;
            }
            string author = prop.Author;
            DateTime creationDate = prop.CreatedDate;
            string modifier = prop.ModifiedBy;
            DateTime modificationDate = prop.ModifiedDate;
            //return prop;
        }
    }

    [TestClass]
    public class ParagraphTests
    {
        private static string testDir = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/";
        private static string processedDir = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/ProcessedDocs/";
        private static string testDoc = "UbuntuSoftwareCenter";

        [TestMethod]
        public void TestParagraphText()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();

            // Act
            para.Text = "This is a paragraph text.";
            body.AppendChild(para);

            // Assert
            var actualText = para.Text;
            var expectedText = "This is a paragraph text.";
            Assert.AreEqual(expectedText, actualText);
        }

        [TestMethod]
        public void TestParagraphAlign()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();

            // Act
            para.Align = "Left";
            body.AppendChild(para);

            // Assert

            var actualAlign = para.Align[0].ToString().ToUpper() + para.Align.Substring(1);
            var expectedAlign = "Left";
            Assert.AreEqual(expectedAlign, actualAlign);
        }

        [TestMethod]
        public void TestParagraphIndent()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();

            // Act
            para.Indent = "700";
            body.AppendChild(para);

            // Assert

            var actualIndent = para.Indent;
            var expectedIndent = "700";
            Assert.AreEqual(expectedIndent, actualIndent);
        }
        [TestMethod]
        public void TestParagraphRightIndent()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();

            // Act
            para.RihgtIndent = "650";
            body.AppendChild(para);

            // Assert

            var actualIndent = para.RihgtIndent;
            var expectedIndent = "650";
            Assert.AreEqual(expectedIndent, actualIndent);
        }
        [TestMethod]
        public void TestParagraphFirstLineIndent()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();

            // Act
            para.FirstLineIndent = "690";
            body.AppendChild(para);

            // Assert

            var actualIndent = para.FirstLineIndent;
            var expectedIndent = "690";
            Assert.AreEqual(expectedIndent, actualIndent);
        }

        [TestMethod]
        public void TestParagraphLineSpacing()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();

            // Act
            para.LinesSpacing = "700";
            body.AppendChild(para);

            // Assert

            var actualLineSpacing = para.LinesSpacing;
            var expectedLineSpacing = "700";
            Assert.AreEqual(expectedLineSpacing, actualLineSpacing);
        }

        [TestMethod]
        public void GetRuns_ReturnsCorrectFormatting()
        {
            // Create a new document and add a paragraph with some runs
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();
            para.Text = "This is some ";
            var run1 = new Run();
            run1.Text = "bold and italic text.";
            run1.Bold = true;
            run1.Italic = true;

            para.AppendChild(run1);
            body.AppendChild(para);
            doc.Save(processedDir + "GetRunsTest_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx");
            // Retrieve the runs from the paragraph
            var runs = para.GetRuns().ToList();

            // Check that the runs have the correct formatting

            //Assert.That(runs.Count, ));
            Assert.AreEqual(2, runs.Count);
            Assert.AreEqual("This is some ", runs[0].Text);
            Assert.IsFalse(runs[0].Bold);
            Assert.IsFalse(runs[0].Italic);
            Assert.AreEqual("bold and italic text.", runs[1].Text);
            Assert.IsTrue(runs[1].Bold);
            Assert.IsTrue(runs[1].Italic);
            Assert.IsFalse(runs[1].Underline);
            Assert.IsNull(runs[1].FontFamily);
            Assert.AreEqual(0, runs[1].FontSize);
            Assert.IsNull(runs[1].Color);
            Assert.AreEqual("This is some bold and italic text.", para.Text);

        }

    }

    [TestClass]
    public class RunTests
    {

        private static string testDir = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/";
        private static string processedDir = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/ProcessedDocs/";
        private static string testDoc = "UbuntuSoftwareCenter";
        [TestMethod]
        public void TestRunText()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();
            var run = new Run();

            // Act
            run.Text = "This is a test run.";
            para.AppendChild(run);
            body.AppendChild(para);
            doc.Save(processedDir + "TestRunText_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx");

            // Assert
            var actualText = run.Text;
            var expectedText = "This is a test run.";
            Assert.AreEqual(expectedText, actualText);
        }

        [TestMethod]
        public void TestRunIsBold()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();
            var run = new Run();

            // Act
            run.Bold = true;
            para.AppendChild(run);
            body.AppendChild(para);
            doc.Save(processedDir + "TestRunIsBold_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx");
            // Assert
            var actualIsBold = run.Bold;
            var expectedIsBold = true;
            Assert.AreEqual(expectedIsBold, actualIsBold);
        }

        [TestMethod]
        public void TestRunIsItalic()
        {
            // Arrange
            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();
            var run = new Run();

            // Act
            run.Italic = true;
            para.AppendChild(run);
            body.AppendChild(para);
            doc.Save(processedDir + "TestTestRunIsItalic_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx");
            // Assert
            var actualIsItalic = run.Italic;
            var expectedIsItalic = true;
            Assert.AreEqual(expectedIsItalic, actualIsItalic);
        }


    }

    [TestClass]
    public class TestsTableClass
    {
        private const string TestFile = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/Docs.docx";

        [TestMethod]
        public void TestCreateTable()
        {
            var doc = new Document();
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
      

            int count = body.FindTableByText("Name");
            Assert.AreEqual(count, 1);
            foreach (TableRow row in body.FindTableRow(0, 1))
            {
                Assert.AreEqual(int.Parse(row.NumberOfCell), 3);
            }
            foreach (TableCell cell in body.FindTableCell(0, 0, 0))
            {
                Assert.AreEqual(cell.CellWidth, "2400");
                Assert.AreEqual(cell.Text, "Name");
            }

            Assert.AreEqual(body.getDocumentTables.Count(), 1);
            foreach (FileFormat.Words.Table.Table props in body.getDocumentTables)
            {
                Assert.AreEqual(int.Parse(props.NumberOfRows), 3);
                Assert.AreEqual(int.Parse(props.NumberOfColumns), 3);
                Assert.AreEqual(int.Parse(props.NumberOfCells), 9);
                Assert.AreEqual(props.TableBorder, "basicBlackSquares");
                Assert.AreEqual(props.CellWidth, "2400");
                Assert.AreEqual(props.TablePosition, "left");

            }
            Assert.AreEqual(table.ChangeTextInCell(TestFile, 0, 0, 0, "changed"), "Cell updated successfully");
            Assert.AreEqual(table.ChangeTextInCell(TestFile, 10, 0, 0, "changed"), "table index out of range");
            Assert.AreEqual(table.ChangeTextInCell(TestFile, 0, 10, 0, "changed"), "table row index out of range");
            Assert.AreEqual(table.ChangeTextInCell(TestFile, 0, 0, 10, "changed"), "table cell index out of range");

        }
    }

    [TestClass]
    public class ImageTests
    {
        private static string testDir = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/";
        private static string processedDir = "/Users/Mustafa/Projects/FileFormat.Words/TestDocs/ProcessedDocs/";
        private static string testDoc = "UbuntuSoftwareCenter";
        [TestMethod]
        public void TestAddImageToDoc()
        {

            // Arrange
            var documentPath = processedDir + "TestRunImage_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx";
            var imagePath = testDir + "testimage.jpeg";


            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();
            var run = new Run();
            var image = new FileFormat.Words.Image(doc, imagePath, 100, 100);

            // Act
            run.AppendChild(image);
            para.AppendChild(run);

            body.AppendChild(para);
            doc.Save(documentPath);

            // Assert
            Assert.IsTrue(File.Exists(documentPath));

        }

        [TestMethod]
        public void TestExtractImagesFromDocument()
        {
            // Arrange
            var documentPath = processedDir + "TestRunImage_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx";
            var imagePath1 = testDir + "testimage.jpeg";
            var imagePath2 = testDir + "testimage2.jpeg";



            var doc = new Document();
            var body = new Body(doc);
            var para = new Paragraph();
            var run = new Run();
            var imagePart1 = new FileFormat.Words.Image(doc, imagePath1, 100, 100);
            var imagePart2 = new FileFormat.Words.Image(doc, imagePath2, 100, 100);

            // Act
            run.AppendChild(imagePart1);
            run.AppendChild(imagePart2);
            para.AppendChild(run);
            body.AppendChild(para);


            var images = FileFormat.Words.Image.ExtractImagesFromDocument(doc);

            int imageCount = images.Count;

            // Assert
            Assert.AreEqual(2, imageCount);
            var i = 1;
            foreach (Stream imagePart in images)
            {
                using (FileStream stream = new FileStream(testDir + "/" + i + ".jpeg", FileMode.Create))
                {
                    imagePart.CopyTo(stream);
                }
                i = i + 1;
            }
            Console.ReadLine();



        }

        [TestMethod]
        public void TestResizeImage()
        {
            // Arrange
            var documentPath = processedDir + "TestResizeImage_" + DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") + ".docx";
            var imagePath = testDir + "testimage.jpeg";
            int initialWidth = 500;
            int initialHeight = 300;
            int newWidth = 400;
            int newHeight = 200;

            var doc = new Document();
            var image = new FileFormat.Words.Image(doc, imagePath, initialWidth, initialHeight);

            // Act
            image.ResizeImage(newWidth, newHeight);

            // Assert

            //newWidth * 9525 is used to convert the newWidth from pixels to EMUs by assuming a conversion of 96 pixels per inch.
            Assert.AreEqual(newWidth * 9525, image.GetExtentCx());
            Assert.AreEqual(newHeight * 9525, image.GetExtentCy());


        }
    }


}