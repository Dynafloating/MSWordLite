using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MSWordLite.Elements;
using System.Collections.Generic;
using System.Linq;

namespace MSWordLite
{
    /// <summary>
    /// Word 範本文件
    /// </summary>
    class Document
    {
        /// <summary>
        /// Word document
        /// </summary>
        public WordprocessingDocument WordDocument { get; set; }

        /// <summary>
        /// Word document root.
        /// </summary>
        public OpenXmlPartRootElement RootElement => WordDocument.MainDocumentPart.RootElement;

        /// <summary>
        /// All tables in tamplate at current time.
        /// </summary>
        public IEnumerable<Table> WordTables => RootElement.Elements()
            .SelectMany(element => Table.SearchFrom(element));

        /// <summary>
        /// Determine if document has any table.
        /// </summary>
        public bool HasTables => WordTables != null && WordTables.Count() > 0;

        /// <summary>
        /// All bookmarks in template.
        /// </summary>
        public Dictionary<string, Bookmark> WordBookmarks { get; set; } = new Dictionary<string, Bookmark>();

        /// <summary>
        /// Determine if document has any bookmark.
        /// </summary>
        public bool HasBookmarks => WordBookmarks != null && WordBookmarks.Count > 0;

        /// <summary>
        /// Initialize an document object.
        /// </summary>
        /// <param name="wordDocument"></param>
        public Document(WordprocessingDocument wordDocument)
        {
            WordDocument = wordDocument;
        }
    }
}
