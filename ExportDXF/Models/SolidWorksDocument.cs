using ExportDXF.Models;

namespace ExportDXF.Services
{
    /// <summary>
    /// Represents a SolidWorks document with essential metadata.
    /// </summary>
    public class SolidWorksDocument
    {
        /// <summary>
        /// The title/name of the document.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// The full file path of the document.
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// The type of document (Part, Assembly, or Drawing).
        /// </summary>
        public DocumentType DocumentType { get; set; }

        /// <summary>
        /// The native SolidWorks document object (ModelDoc2, PartDoc, etc.).
        /// </summary>
        public object NativeDocument { get; set; }
    }
}