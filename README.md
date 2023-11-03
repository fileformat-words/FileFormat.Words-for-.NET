# Word Document C# API System Outlines

This documentation provides an in-depth overview of the internal namespaces and classes, unveiling the intricacies behind our Word Document C# API system. While the internal structure is complex, we've designed the public APIs to be straightforward and user-friendly, making Word document manipulation a seamless experience.

For a more detailed understanding of our system architecture, design patterns, and public interfaces, please visit the [Articles Section](https://fileformat-words.github.io/FileFormat.Words-for-.NET/).

## FileFormat.Words Namespace

### Document Class
- The primary interface for creating, loading, and modifying Word documents.
- Serves as a facade for interacting with internal classes in the OpenXML.Words, OpenXML.Words.Data, and OpenXML.Templates namespaces.

### FileFormat.Words.IElements (Custom Objects)
- Custom elements, such as Paragraphs, Images, and Tables, that mimic Word document structure.
- Act as a data structure for seamless data transfer to and from OpenXML objects.
- Offers a user-friendly interface for interacting with Word document content.

## OpenXML.Words Namespace

### Document Class (Internal)
- Facilitates loading existing Word documents into OpenXML and creating new documents from scratch.
- Acts as a bridge between custom document elements in FileFormat.Words and OpenXML-based Word documents.
- Sets the WordProcessing package for OpenXML.Words.Data.OOXMLDocData to enable synchronization.
- Utilizes templates from the OpenXML.Templates namespace for creating new documents.

## OpenXML.Words.Data Namespace

### OOXMLDocData Class (Internal, Static)
- Employs static operations for inserting, updating, and removing elements in an OpenXML-based Word document.
- Receives synchronization instructions from the FileFormat.Words.Document class.
- Guarantees changes made to custom objects are accurately reflected in the OpenXML document.

## OpenXML.Templates Namespace

- Offers pre-defined templates created using OpenXML SDK Productivity Tools.
- Comprises classes for core properties and custom properties, enhancing document metadata and customization.
- Templates come into play when creating new Word documents from scratch within the OpenXML.Words.Document class.

This extended architecture leverages templates and properties from the OpenXML.Templates namespace, enriching the document creation process with pre-defined structures and metadata customization options for newly created documents. The FileFormat.Words.Document class remains the central interface for users, orchestrating interactions with various internal components across multiple namespaces.

## API Reference
- [API Reference](https://fileformat-words.github.io/FileFormat.Words-for-.NET/api/index.html) - In-depth information about public interfaces and usage.

## Technical Docs
- [Articles](https://fileformat-words.github.io/FileFormat.Words-for-.NET/articles/intro.html) - Comprehensive insights into the system architecture, design patterns, and API usage in different scenarios.

# Installation
- Install-Package FileFormat.Words

# System Requirements
- .NET Core 3.1 and above
