using System;
using System.Collections.Generic;
using IO = System.IO;
using OOXML = OpenXML.Words;
using OWD = OpenXML.Words.Data;
using FileFormat.Words.IElements;
using System.Threading;

namespace FileFormat.Words
{
    /// <summary>
    /// Custom exception class for file format-related exceptions.
    /// </summary>
    public class FileFormatException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="FileFormatException"/> class with a specified error message and a reference to the inner exception.
        /// </summary>
        /// <param name="message">The error message that explains the reason for the exception.</param>
        /// <param name="innerException">The exception that is the cause of the current exception, or a null reference if no inner exception is specified.</param>
        public FileFormatException(string message, Exception innerException) : base(message, innerException)
        {
            //Do nothing
        }
    }
    /// <summary>
    /// Represents a document with structured elements.
    /// </summary>
    public class Document : IDisposable
    {
        private List<IElement> _lstStructure;
        private ElementStyles _elementStyles;// { get; internal set; }
        private OOXML.OwDocument _ooxmlDoc;
        private bool _isNew = true;
        //private static List<int> _instances = new List<int>();
        //private static int _instanceNumber = 0;
        //private int _instance = 0;
        private static Lazy<List<int>> _instances = new Lazy<List<int>>(() => new List<int>());
        private static int _instanceNumber = 0;
        private int _instance = 0;
        private int _originalSize = 0;
        private IO.MemoryStream _ms; //= new IO.MemoryStream();
        private readonly object _lockObject = new object(); // Lock object
        /// <summary>
        /// Initializes a new instance of the <see cref="Document"/> class.
        /// This constructor creates a new, empty document.
        /// </summary>
        /// <remarks>
        /// Use this constructor to create a new, blank document that you can populate with content.
        /// To work with an existing document, consider using the <see cref="Document(string)"/> constructor.
        /// </remarks>
        /// <example>
        /// <code>
        /// // Create a new, empty document
        /// Document Doc = new Document();
        /// // Initialize a new instance of body the empty document.
        /// Body body = new Body(doc);
        /// // Add paragraph wiht run
        /// Paragraph para = new Paragraph();
        /// para.AddRun(new FF.Run
        ///    {
        ///         Text = "First Run with Times New Roman Blue Color Bold",
        ///         FontFamily = "Times New Roman",
        ///         Color = FF.Colors.Blue,
        ///         Bold = true,
        ///         FontSize = 12
        ///     });
        /// // Append paragraph to the document
        /// body.AppendChild(para);
        /// // Save document to file.
        /// doc.Save("DocumentWithPara.docx");
        /// </code>
        /// </example>
        public Document()
        {
            lock (_lockObject)
            {
                try
                {
                    _isNew = true;
                    _ms = new IO.MemoryStream();
                    _lstStructure = new List<IElement>();
                    _ooxmlDoc = OOXML.OwDocument.CreateInstance();
                    _elementStyles = _ooxmlDoc.LoadStyles();
                }
                catch (Exception ex)
                {
                    var errorMessage = ConstructMessage(ex, "Initializing OWDocument");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="Document"/> class by loading content from a file specified by its filename.
        /// </summary>
        /// <param name="filename">The path to the file from which to load the document content.</param>
        /// <exception cref="FileFormatException">
        /// Thrown if an error occurs while loading the document.
        /// </exception>
        public Document(string filename)
        {
            lock (_lockObject)
            {
                try
                {
                    _isNew = false;
                    _instance = Interlocked.Increment(ref _instanceNumber);
                    _instances.Value.Add(_instance);
                    _ms = new IO.MemoryStream();
                    using (var fs = new IO.FileStream(filename, IO.FileMode.Open))
                    {
                        fs.CopyTo(_ms);
                    }

                    _ooxmlDoc = OOXML.OwDocument.CreateInstance();
                    _lstStructure = _ooxmlDoc.LoadDocument(_ms);
                    _elementStyles = _ooxmlDoc.LoadStyles();
                    _originalSize = _lstStructure.Count;
                }
                catch (Exception ex)
                {
                    var errorMessage = ConstructMessage(ex, "Loading Document from file");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="Document"/> class by loading content from a <see cref="System.IO.Stream"/>.
        /// </summary>
        /// <param name="stream">The input stream from which to load the document content.</param>
        /// <exception cref="FileFormatException">
        /// Thrown if an error occurs while loading the document.
        /// </exception>
        public Document(IO.Stream stream)
        {
            lock (_lockObject)
            {
                try
                {
                    _isNew = false;
                    _instance = Interlocked.Increment(ref _instanceNumber);
                    _instances.Value.Add(_instance);
                    _ms = new IO.MemoryStream();
                    stream.CopyTo(_ms);
                    _ooxmlDoc = OOXML.OwDocument.CreateInstance();
                    _lstStructure = _ooxmlDoc.LoadDocument(_ms);
                    _elementStyles = _ooxmlDoc.LoadStyles();
                    _originalSize = _lstStructure.Count;
                }
                catch (Exception ex)
                {
                    var errorMessage = ConstructMessage(ex, "Loading Document from stream");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Saves the document to a file specified by its filename.
        /// </summary>
        /// <param name="filename">The path to the file where the document will be saved.</param>
        /// <exception cref="FileFormatException">
        /// Thrown if an error occurs while saving the document.
        /// </exception>
        public void Save(string filename)
        {
            lock (_lockObject)
            {
                try
                {
                    if (!_isNew)
                    {
                        using var fs = new IO.FileStream(filename, IO.FileMode.Create);
                        OWD.OoxmlDocData.CreateInstance().Save(fs, this);
                    }
                    else
                    {
                        _ooxmlDoc = OOXML.OwDocument.CreateInstance();
                        _ooxmlDoc.CreateDocument(_lstStructure);
                        using var fs = new IO.FileStream(filename, IO.FileMode.Create);
                        _ooxmlDoc.SaveDocument(fs);
                    }
                }
                catch (Exception ex)
                {
                    var errorMessage = ConstructMessage(ex, "Saving Document to file");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }

        /// <summary>
        /// Saves the document to the specified <see cref="System.IO.Stream"/>.
        /// </summary>
        /// <param name="stream">The stream to which the document will be saved.</param>
        /// <exception cref="FileFormatException">
        /// Thrown if an error occurs while saving the document.
        /// </exception>
        public void Save(IO.Stream stream)
        {
            lock (_lockObject)
            {
                try
                {
                    if (!_isNew)
                    {
                        OWD.OoxmlDocData.CreateInstance().Save(stream, this);
                    }
                    else
                    {
                        _ooxmlDoc = OOXML.OwDocument.CreateInstance();
                        _ooxmlDoc.CreateDocument(_lstStructure);
                        _ooxmlDoc.SaveDocument(stream);
                    }
                }
                catch (Exception ex)
                {
                    var errorMessage = ConstructMessage(ex, "Saving Document to stream");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        private string ConstructMessage(Exception ex, string operation)
        {
            return $"Error in Operation {operation} at FileFormat.Words: {ex.Message} \n Inner Exception: {ex.InnerException?.Message ?? "N/A"}";
        }

        internal int GetInstanceInfo()
        {
            return _instance;
        }

        /// <summary>
        /// Retrieves a list of existing elements from the internal data structure.
        /// </summary>
        /// <returns>
        /// A list of <see cref="IElement"/> objects representing the existing elements.
        /// </returns>
        /// <seealso cref="IElement"/>
        public List<IElement> GetElements()
        {
            return _lstStructure;
        }
        /// <summary>
        /// Gets the element styles of the document.
        /// </summary>
        public ElementStyles GetElementStyles()
        {
            return _elementStyles;
        }
        /// <summary>
        /// Updates an existing element in the structure.
        /// </summary>
        /// <param name="element">The updated element to replace the existing one.</param>
        /// <returns>
        ///   <c>true</c> if the element is successfully updated; otherwise, <c>false</c>.
        ///   If the element is not found in the structure, <c>false</c> is returned.
        ///   If an error occurs during the update operation, a <see cref="FileFormatException"/> is thrown.
        /// </returns>
        /// <remarks>
        /// This method updates an existing element within the structure based on the element's unique identifier.
        /// It locates the element in the structure and replaces it with the updated element.
        /// If the update operation is successful, the method returns <c>true></c>.
        /// If the element is not found in the structure, <c>false</c> is returned, and no changes are made to the structure.
        /// If an exception occurs during the update operation, the method attempts to restore the structure to its previous state
        /// and throws a <see cref="FileFormatException"/> with detailed error information.
        /// </remarks>
        public bool Update(IElement element)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == element.ElementId);
                if (position < 0)
                {
                    return false;
                }

                var backupElement = element;
                try
                {
                    _lstStructure[position] = element;
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Update(element, position, this);
                    return true;
                }
                catch (Exception ex)
                {
                    _lstStructure[position] = backupElement;
                    var errorMessage = ConstructMessage(ex, "Update");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Updates an existing element in the structure.
        /// </summary>
        /// <param name="elementId">The unique identifier of the element to be updated.</param>
        /// <param name="element">The updated element to replace the existing one.</param>
        /// <returns>
        ///   <c>true</c> if the element is successfully updated; otherwise, <c>false</c>.
        ///   If the element with the specified ID is not found in the structure, <c>false</c> is returned.
        ///   If the elements' IDs do not match, a <see cref="FileFormatException"/> is thrown.
        ///   If an error occurs during the update operation, a <see cref="FileFormatException"/> is thrown.
        /// </returns>
        /// <remarks>
        /// This method updates an existing element within the structure based on the provided element's unique identifier.
        /// It verifies that the element to be updated has the same ID as the specified element ID.
        /// If the IDs match, the method attempts to locate the element in the structure and replace it with the updated element.
        /// If the update operation is successful, the method returns <c>true></c>.
        /// If the element with the specified ID is not found, <c>false</c> is returned, and no changes are made to the structure.
        /// If the provided element's ID does not match the specified element ID, a <see cref="FileFormatException"/> is thrown.
        /// If an exception occurs during the update operation, the method attempts to restore the structure to its previous state
        /// and throws a <see cref="FileFormatException"/> with detailed error information.
        /// </remarks>
        public bool Update(int elementId, IElement element)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                if (elementId != element.ElementId)
                {
                    var ex = new Exception("The elements mismatch: Update is only available for same element");
                    throw new FileFormatException(ex.Message, new InvalidOperationException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == elementId);
                if (position < 0)
                {
                    return false;
                }

                var backupElement = element;
                try
                {
                    _lstStructure[position] = element;
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Update(element, position, this);
                    return true;
                }
                catch (Exception ex)
                {
                    _lstStructure[position] = backupElement;
                    var errorMessage = ConstructMessage(ex, "Update");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Inserts an element before the specified reference element.
        /// </summary>
        /// <param name="newElement">The element to be inserted.</param>
        /// <param name="element">The reference element before which the new element should be inserted.</param>
        /// <returns>
        ///   The unique identifier <see cref="IElement.ElementId"/> of inserted element if the new element is successfully inserted before the reference element.
        ///   If the reference element is not found in the structure, -1 is returned.
        ///   If an error occurs during the insertion operation, a <see cref="FileFormatException"/> is thrown.
        /// </returns>
        /// <remarks>
        /// This method inserts the specified element before the provided reference element in the structure.
        /// The method attempts to locate the reference element and insert the new element before it.
        /// If the insertion operation is successful, the method returns <see cref="IElement.ElementId"/> of the inserted element.
        /// If the reference element is not found, -1 is returned, and no changes are made to the structure.
        /// If an exception occurs during the insertion operation, the method attempts to restore the structure to its previous state
        /// and throws a <see cref="FileFormatException"/> with detailed error information.
        /// </remarks>
        public int InsertBefore(IElement newElement, IElement element)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == element.ElementId);
                if (position < 0)
                {
                    return -1;
                }

                var newElementId = _originalSize + 1;

                switch (newElement)
                {
                    case Paragraph para:
                        para.ElementId = newElementId;
                        break;
                    case Table table:
                        table.ElementId = newElementId;
                        break;
                    case Image image:
                        image.ElementId = newElementId;
                        break;
                }

                try
                {
                    _lstStructure.Insert(position, newElement);
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Insert(newElement, position, this);
                    _originalSize++;
                    return newElementId;
                }
                catch (Exception ex)
                {
                    _lstStructure.RemoveAt(position);
                    var errorMessage = ConstructMessage(ex, "InsertBefore");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Inserts an element before the element with the specified unique ID.
        /// </summary>
        /// <param name="newElement">The element to be inserted.</param>
        /// <param name="elementId">The unique ID of the element before which the new element should be inserted.</param>
        /// <returns>
        ///   The unique identifier <see cref="IElement.ElementId"/> of the inserted element if the new element is successfully inserted before the specified element.
        ///   If the specified element with the provided ID is not found, -1 is returned.
        ///   If an error occurs during the insertion operation, a <see cref="FileFormatException"/> is thrown.
        /// </returns>
        /// <remarks>
        /// This method inserts the specified element before the element with the provided unique ID in the structure.
        /// The method attempts to locate the element with the specified ID and insert the new element before it.
        /// If the insertion operation is successful, the method returns <see cref="IElement.ElementId"/> of the inserted element.
        /// If the element with the specified ID is not found, -1 is returned, and no changes are made to the structure.
        /// If an exception occurs during the insertion operation, the method attempts to restore the structure to its previous state
        /// and throws a <see cref="FileFormatException"/> with detailed error information.
        /// </remarks>
        public int InsertBefore(IElement newElement, int elementId)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == elementId);
                if (position < 0) return -1;
                var newElementId = _originalSize + 1;
                switch (newElement)
                {
                    case Paragraph para:
                        para.ElementId = newElementId;
                        break;
                    case Table table:
                        table.ElementId = newElementId;
                        break;
                    case Image image:
                        image.ElementId = newElementId;
                        break;
                }

                try
                {
                    _lstStructure.Insert(position, newElement);
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Insert(newElement, position, this);
                    _originalSize++;
                    return newElementId;
                }
                catch (Exception ex)
                {
                    _lstStructure.RemoveAt(position);
                    var errorMessage = ConstructMessage(ex, "InsertBefore");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        internal int Append(IElement newElement)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                //Console.WriteLine($"Number of Elements in List: {_lstStructure.Count}");
                var newElementId = _originalSize + 1;
                switch (newElement)
                {
                    //Console.WriteLine(_lstStructure.Count);
                    case Paragraph para:
                        para.ElementId = newElementId;
                        break;
                    case Table table:
                        table.ElementId = newElementId;
                        break;
                    case Image image:
                        image.ElementId = newElementId;
                        break;
                }

                var originalCount = _lstStructure.Count;
                try
                {
                    if (_lstStructure.Count > 0 && _lstStructure[originalCount - 1] is Section section)
                    {
                        InsertBefore(newElement, section.ElementId);
                    }
                    else
                    {
                        _lstStructure.Add(newElement);
                        if (!_isNew) OWD.OoxmlDocData.CreateInstance().Append(newElement, this);
                    }

                    //Console.WriteLine("Hello + " + _instance);
                    _originalSize++;
                    return newElementId;
                }
                catch (Exception ex)
                {
                    // Determine which element caused the exception and remove it
                    if (_lstStructure.Count > originalCount)
                    {
                        if (_lstStructure[originalCount - 1] is Section section)
                        {
                            RemoveBefore(section.ElementId);
                        }
                        else
                        {
                            _lstStructure.RemoveAt(originalCount - 1); // Corrected index
                        }
                    }

                    var errorMessage = ConstructMessage(ex, "Append");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Inserts an element after a specified element in the structure.
        /// </summary>
        /// <param name="newElement">The element to be inserted.</param>
        /// <param name="element">The element after which the new element should be inserted.</param>
        /// <returns>
        ///   The unique identifier <see cref="IElement.ElementId"/> if the new element is successfully inserted after the specified element.
        ///   If the specified element is not found, -1 is returned.
        ///   If an error occurs during the insertion operation, a <see cref="FileFormatException"/> is thrown.
        /// </returns>
        /// <remarks>
        /// This method inserts the specified element after the provided element in the structure.
        /// The method attempts to locate the specified element by comparing their unique IDs.
        /// If the insertion operation is successful, the method returns <see cref="IElement.ElementId"/> of the inserted element.
        /// If the specified element is not found, -1 is returned, and no changes are made to the structure.
        /// If an exception occurs during the insertion operation, the method attempts to restore the structure to its previous state
        /// and throws a <see cref="FileFormatException"/> with detailed error information.
        /// </remarks>
        public int InsertAfter(IElement newElement, IElement element)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == element.ElementId);
                if (position < 0) return -1;
                var newElementId = _originalSize + 1;
                switch (newElement)
                {
                    case Paragraph para:
                        para.ElementId = newElementId;
                        break;
                    case Table table:
                        table.ElementId = newElementId;
                        break;
                    case Image image:
                        image.ElementId = newElementId;
                        break;
                }

                try
                {
                    _lstStructure.Insert(position + 1, newElement);
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Insert(newElement, position + 1, this);
                    _originalSize++;
                    return newElementId;
                }
                catch (Exception ex)
                {
                    _lstStructure.RemoveAt(position + 1);
                    var errorMessage = ConstructMessage(ex, "InsertAfter");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Inserts an element after a specified element in the structure by its unique ID.
        /// </summary>
        /// <param name="newElement">The element to be inserted.</param>
        /// <param name="elementId">The unique ID of the element after which the new element should be inserted.</param>
        /// <returns>
        ///   The unique identifier <see cref="IElement.ElementId"/> of the inserted element if the new element is successfully inserted after the specified element.
        ///   If the specified element is not found, -1 is returned.
        ///   If an error occurs during the insertion operation, a <see cref="FileFormatException"/> is thrown.
        /// </returns>
        /// <remarks>
        /// This method inserts the specified element after the element with the provided unique ID in the structure.
        /// It is important to ensure that the specified element exists in the structure to determine the insertion point.
        /// If the insertion operation is successful, the method returns <see cref="IElement.ElementId"/> of the inserted element.
        /// If the specified element is not found, -1 is returned, and no changes are made to the structure.
        /// If an exception occurs during the insertion operation, the method attempts to restore the structure to its previous state
        /// and throws a <see cref="FileFormatException"/> with detailed error information.
        /// </remarks>
        public int InsertAfter(IElement newElement, int elementId)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }
                var position = _lstStructure.FindIndex(e => e.ElementId == elementId);
                if (position < 0) return -1;
                var newElementId = _originalSize + 1;
                switch (newElement)
                {
                    case Paragraph para:
                        para.ElementId = newElementId;
                        break;
                    case Table table:
                        table.ElementId = newElementId;
                        break;
                    case Image image:
                        image.ElementId = newElementId;
                        break;
                }

                try
                {
                    _lstStructure.Insert(position + 1, newElement);
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Insert(newElement, position + 1, this);
                    _originalSize++;
                    return newElementId;
                }
                catch (Exception ex)
                {
                    _lstStructure.RemoveAt(position + 1);
                    var errorMessage = ConstructMessage(ex, "InsertAfter");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Removes the element that follows a specified element in the structure.
        /// </summary>
        /// <param name="element">The element whose follower should be removed.</param>
        /// <returns>
        ///   <c>true</c> if the element following the specified element is successfully removed; otherwise, <c>false</c>.
        ///   If the specified element is not found or if there is no element following it, <c>false</c> is returned.
        ///   If an error occurs during the removal operation, a <see cref="FileFormatException"/> is thrown.
        /// </returns>
        /// <remarks>
        /// This method removes the element that comes after the specified element in the structure. It is essential to ensure that the
        /// specified element exists in the structure and has an element following it.
        /// If the removal operation is successful, the method returns <c>true</c>.
        /// If the element is not found or there is no element following it, <c>false</c> is returned, and no changes are made to the structure.
        /// If an exception occurs during the removal operation, the method attempts to restore the structure to its previous state
        /// and throws a <see cref="FileFormatException"/> with detailed error information.
        /// </remarks>
        public bool RemoveAfter(IElement element)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }
                var position = _lstStructure.FindIndex(e => e.ElementId == element.ElementId);
                if (position >= 0 && position < _lstStructure.Count - 1)
                {
                    var backupElement = _lstStructure[position + 1];
                    try
                    {
                        _lstStructure.RemoveAt(position + 1);
                        if (!_isNew) OWD.OoxmlDocData.CreateInstance().Remove(position + 1, this);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        _lstStructure.Insert(position + 1, backupElement);
                        var errorMessage = ConstructMessage(ex, "RemoveAfter");
                        throw new FileFormatException(errorMessage, ex);
                    }
                }

                return false;
            }
        }
        /// <summary>
        /// Removes the element that follows a specified element in the structure by its unique ID.
        /// </summary>
        /// <param name="elementId">The unique ID of the element whose follower should be removed.</param>
        /// <returns>
        ///   <c>true</c> if the element following the specified element is successfully removed; otherwise, <c>false</c>.
        ///   If the specified element is not found or if there is no element following it, <c>false</c> is returned.
        ///   If an error occurs during the removal operation, a <see cref="FileFormatException"/> is thrown.
        /// </returns>
        /// <remarks>
        /// This method removes the element that comes after the specified element with the provided unique ID in the structure.
        /// It is essential to ensure that the specified element exists in the structure and has an element following it.
        /// If the removal operation is successful, the method returns <c>true</c>.
        /// If the element is not found or there is no element following it, <c>false</c> is returned, and no changes are made to the structure.
        /// If an exception occurs during the removal operation, the method attempts to restore the structure to its previous state
        /// and throws a <see cref="FileFormatException"/> with detailed error information.
        /// </remarks>
        public bool RemoveAfter(int elementId)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == elementId);
                if (position >= 0 && position < _lstStructure.Count - 1)
                {
                    var backupElement = _lstStructure[position + 1];
                    try
                    {
                        _lstStructure.RemoveAt(position + 1);
                        if (!_isNew) OWD.OoxmlDocData.CreateInstance().Remove(position + 1, this);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        _lstStructure.Insert(position + 1, backupElement);
                        var errorMessage = ConstructMessage(ex, "RemoveAfter");
                        throw new FileFormatException(errorMessage, ex);
                    }
                }

                return false;
            }
        }
        /// <summary>
        /// Removes the element that precedes the specified <paramref name="element"/> from the collection.
        /// If the specified <paramref name="element"/> is not found or is the first element in the collection, no removal occurs.
        /// </summary>
        /// <remarks>
        /// The method searches for the element with the specified <paramref name="element"/> identifier and removes
        /// the element that immediately precedes it in the collection. If the specified <paramref name="element"/>
        /// is not found or is the first element in the collection, no removal occurs, and the method returns -1.
        /// In case of success, it returns the <see cref="IElement.ElementId"/> of removed element.
        /// If the removal of the preceding element fails due to an exception, it rolls back the operation and reverts the document to its original state.
        /// </remarks>
        /// <param name="element">The element whose predecessor should be removed.</param>
        /// <returns>
        ///   <c>true</c> if the preceding element is successfully removed; otherwise, <c>false</c> if the element is not found.
        ///   Throws a <see cref="FileFormatException"/> if an exception occurs during the operation.
        /// </returns>
        /// <seealso cref="IElement.ElementId"/>
        public bool RemoveBefore(IElement element)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == element.ElementId);
                if (position <= 0) return false;
                var backupElement = element;
                try
                {
                    _lstStructure.RemoveAt(position - 1);
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Remove(position - 1, this);
                    return true;
                }
                catch (Exception ex)
                {
                    _lstStructure.Insert(position - 1, backupElement);
                    var errorMessage = ConstructMessage(ex, "RemoveBefore");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Removes the element that precedes the specified element with the given ElementID in the document.
        /// </summary>
        /// <param name="elementId">The unique identifier of the element after which you want to remove the preceding element.</param>
        /// <returns>
        ///   <c>true</c> if the preceding element is successfully removed; otherwise, <c>false</c> if the element is not found.
        ///   Throws a <see cref="FileFormatException"/> if an exception occurs during the operation.
        /// </returns>
        /// <remarks>
        /// If the specified element is not found in the document, this method returns <c>false</c>.
        /// If the removal of the preceding element fails due to an exception, it rolls back the operation and reverts the document to its original state.
        /// The preceding element is removed from the internal document structure and, if applicable, the underlying OOXML document.
        /// </remarks>
        public bool RemoveBefore(int elementId)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == elementId);
                if (position <= 0) return false;
                var backupElement = _lstStructure[position];
                try
                {
                    _lstStructure.RemoveAt(position - 1);
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Remove(position - 1, this);
                    return true;
                }
                catch (Exception ex)
                {
                    _lstStructure.Insert(position - 1, backupElement);
                    var errorMessage = ConstructMessage(ex, "RemoveBefore");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Removes the specified element with the given ElementID from the document.
        /// </summary>
        /// <param name="elementId">The unique identifier of the element you want to remove.</param>
        /// <returns>
        ///   <c>true</c> if the element is successfully removed; otherwise, <c>false</c> if the element is not found.
        ///   Throws a <see cref="FileFormatException"/> if an exception occurs during the operation.
        /// </returns>
        /// <remarks>
        /// If the specified element is not found in the document, this method returns <c>false</c>.
        /// If the removal of the element fails due to an exception, it rolls back the operation and reverts the document to its original state.
        /// The element is removed from the internal document structure and, if applicable, the underlying OOXML document.
        /// </remarks>
        public bool Remove(int elementId)
        {
            lock (_lockObject)
            {
                if (_lstStructure == null)
                {
                    throw new FileFormatException("Structure is unavailable...", new NullReferenceException());
                }

                var position = _lstStructure.FindIndex(e => e.ElementId == elementId);
                if (position < 0) return false;
                var backupElement = _lstStructure[position];
                try
                {
                    _lstStructure.RemoveAt(position);
                    if (!_isNew) OWD.OoxmlDocData.CreateInstance().Remove(position, this);
                    return true;
                }
                catch (Exception ex)
                {
                    _lstStructure.Insert(position, backupElement);
                    var errorMessage = ConstructMessage(ex, "Remove");
                    throw new FileFormatException(errorMessage, ex);
                }
            }
        }
        /// <summary>
        /// Dispose off all managed and unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose of managed resources (if any)
                if (_lstStructure != null) _lstStructure = null;
                if (_elementStyles != null) _elementStyles = null;
                if (_ooxmlDoc != null)
                {
                    _ooxmlDoc.Dispose();
                    _ooxmlDoc = null;
                }
                OWD.OoxmlDocData.CreateInstance().Dispose();
            }
            // Dispose of unmanaged resources
            if (_ms == null) return;
            _ms.Dispose();
            _ms = null;
        }
    }
    /// <summary>
    /// Represents the body of a document, containing paragraphs, tables, images, and sections.
    /// </summary>
    public class Body
    {
        private readonly object _lockObject = new object();
        /// <summary>
        /// Gets the list of paragraphs in the body.
        /// </summary>
        public List<Paragraph> Paragraphs { get; internal set; }
        /// <summary>
        /// Gets the list of tables in the body.
        /// </summary>
        public List<Table> Tables { get; internal set; }
        /// <summary>
        /// Gets the list of images in the body.
        /// </summary>
        public List<Image> Images { get; internal set; }
        /// <summary>
        /// Gets the list of sections in the body.
        /// </summary>
        public List<Section> Sections { get; internal set; }
        private Document Doc { get; set; }
        /// <summary>
        /// Initializes a new instance of the <see cref="Body"/> class with the specified document.
        /// </summary>
        /// <param name="doc">The parent document containing the body.</param>
        public Body(Document doc)
        {
            lock (_lockObject)
            {
                Paragraphs = new List<Paragraph>();
                Tables = new List<Table>();
                Images = new List<Image>();
                Sections = new List<Section>();
                foreach (var element in doc.GetElements())
                {
                    if (element is Paragraph)
                    {
                        Paragraphs.Add((Paragraph)element);
                    }

                    if (element is Table)
                    {
                        Tables.Add((Table)element);
                    }

                    if (element is Image)
                    {
                        Images.Add((Image)element);
                    }

                    if (element is Section section)
                    {
                        Sections.Add(section);
                    }
                }

                Doc = doc;
            }
        }
        /// <summary>
        /// Appends a child element to the body.
        /// </summary>
        /// <param name="element">The element to append to the body.</param>
        /// <returns>
        /// The unique identifier <see cref="IElement.ElementId"/> of the appended element.
        /// </returns>
        public int AppendChild(IElement element)
        {
            lock (_lockObject)
            {
                return Doc.Append(element);
            }
        }
    }
}

