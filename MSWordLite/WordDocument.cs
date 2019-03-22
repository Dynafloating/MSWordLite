using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using MSWordLite.Elements;
using System.Collections.Generic;

namespace MSWordLite
{
    /// <summary>
    /// Word 範本文件
    /// </summary>
    class Document
    {
        /// <summary>
        /// Word 文件
        /// </summary>
        public WordprocessingDocument WordDocument { get; set; }

        /// <summary>
        /// Word 文件根
        /// </summary>
        public OpenXmlPartRootElement RootElement => WordDocument.MainDocumentPart.RootElement;

        /// <summary>
        /// 範本中所有的表格
        /// </summary>
        public List<Table> WordTables { get; set; } = new List<Table>();

        public bool HasTables => WordTables != null && WordTables.Count > 0;

        /// <summary>
        /// 範本中所有的書籤
        /// </summary>
        public Dictionary<string, Bookmark> WordBookmarks { get; set; } = new Dictionary<string, Bookmark>();

        public bool HasBookmarks => WordBookmarks != null && WordBookmarks.Count > 0;

        public Document(WordprocessingDocument wordDocument)
        {
            WordDocument = wordDocument;
        }
    }
}
