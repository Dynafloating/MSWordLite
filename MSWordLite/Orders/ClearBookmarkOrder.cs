using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MSWordLite.Orders
{
    /// <summary>
    /// Remove bookmarks in document
    /// </summary>
    public class ClearBookmarkOrder : IOrder
    {
        /// <summary>
        /// If this order is valid.
        /// </summary>
        public bool Valid => Names != null || Regex != null;

        public List<string> Names { get; set; } = new List<string>();

        public Regex Regex { get; set; }

        /// <summary>
        /// Create an clear order to remove all bookmarks in doucment.
        /// </summary>
        public ClearBookmarkOrder() { }

        /// <summary>
        /// Create an clear order to remove specific name of bookmarks in document.
        /// </summary>
        /// <param name="names"></param>
        public ClearBookmarkOrder(List<string> names)
        {
            Names = names;
        }

        /// <summary>
        /// Create an clear order and use regex pattern to test if bookmarks need to be remove in document.
        /// </summary>
        /// <param name="regexPattern"></param>
        public ClearBookmarkOrder(string regexPattern)
        {
            Regex = new Regex(regexPattern);
        }
    }
}
