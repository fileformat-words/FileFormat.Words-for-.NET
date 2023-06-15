# Open-Source .NET Library For Word Document Automation

<p> <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/">FileFormat.Words</a> is an <b>Open-Source</b> .NET API to create customized Microsoft <a href="https://docs.fileformat.com/word-processing/docx/">Word<a/> documents programmatically. This C# library is easy to install, robust, lightweight and offers a
wide range of features to create and manipulate Word documents using a few lines of source code. FileFormat.Words is written in C# and is based on <a href="https://learn.microsoft.com/en-us/office/open-xml/word-processing">OpenXML</a> that <a href="https://www.microsoft.com/">Microsoft</a> backs.
FileFormat.Words is a wrapper that makes use of OpenXML SDK immensely and enables developers to use its features easily.
This <b>Open-Source .NET library</b> helps developers automate Word document generation & manipulation without depending upon any third-party library.
</p>

## About this Repo

<table>
  <tr>
    <th>Directory</th>
    <th>Description</th>
  </tr>
  <tr>
    <td><a href = "https://github.com/fileformat-words/FileFormat.Words-for-.NET/tree/main/FileFormat.Words-Tests">FileFormat.Words-Tests</a></td>
    <td>This directory contains the unit tests of all the features FileFormat.Words offers.</td>
  </tr>
  <tr>
    <td><a href = "https://github.com/fileformat-words/FileFormat.Words-for-.NET/tree/main/FileFormat.Words">FileFormat.Words</a></td>
    <td>It contains all the source code files necessary to execute the functionalities.</td>
  </tr>
  <tr>
    <td><a href = "https://github.com/fileformat-words/FileFormat.Words-for-.NET/tree/main/TestDocs">TestDocs</a></td>
    <td>This folder contains test documents generated by this Open-Souorce .NET library.</td>
  </tr>
</table>

## Library Features & Provisions

<p> <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/api/index.html">FileFormat.Words</a> not only offers provisions to create new <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/api/FileFormat.Words.html">Word documents</a> but also lets developers read & modify
the existing documents programmatically. This .NET library is enterprise-level and  all the processes happen seamlessly.</p>

This Open-Source .NET library offers the following features:

 - <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/">FileFormat.Words</a> offers empty Word document creation as well as with the content. Developers can
   open existing Word document from a file & stream.
 - This Open-Source API lets you add <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/api/FileFormat.Words.Paragraph.html">Paragraphs</a> to the document. Developers can make the text Bold, Italic, and can set the various props such as alignment, style and more.
 - Developers can add <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/api/FileFormat.Words.Table.Table.html">Tables</a> using this Open-Source .NET API. There are many features offered by the <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/api/FileFormat.Words.Table.html">Table namespace</a> such as creating tables, setting table border style, and setting table width.
   In addition, developers can read tables along with all the props(i.e. rows, columns and more) from an existing Word document, and can add/update/remove rows and cells. Further, it lets you edit text inside table cells & more.
 - The <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/api/FileFormat.Words.Image.html">Image</a> class lets you add images with custom props into a Word document. 

## Getting Started With FileFormat.Words For .NET

<p> The installation procedure of this Open-Source <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/api/index.html">.NET API</a> is just one command away as it is also available as a <a href="https://www.nuget.org/profiles/fileformatcom">Nuget Package</a>. Simply, run the following command in the NuGet Package Manager and you are all set to leverage this document
automation library.</p>
<code>Install-Package FileFormat.Words</code>

## Creating a Word Document Programmatically

The following code snippet creates an empty <a href="https://docs.fileformat.com/word-processing/docx/">Word<a/> document programmatically. 
<pre>
<code>
// Create an instance of the Document class.
Document doc = new Document();

// Invoke the Save method to save the Word document onto the disk.
doc.Save("/Docs.docx");
</code>
</pre>

## Coming updates
<p> <a href="https://fileformat-words.github.io/FileFormat.Words-for-.NET/">FileFormat.Words</a> is intended to add more features to its stack. Further, the development of FileFormat.Cells and FileFormat.Slides are a work in progress. So, stay tuned for the upcoming updates. </p>