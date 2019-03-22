using MSWordLite.Elements;
using System.Collections.Generic;
using System.Linq;
using Table = MSWordLite.Elements.Table;

namespace MSWordLite.Processes
{
    class Initializer
    {
        public static void Bookmarks(Document document)
        {
            if (!document.HasBookmarks)
            {
                var bookmarkMap = new Dictionary<string, Bookmark>();
                foreach (var element in document.RootElement.Elements())
                {
                    Bookmark.SearchFrom(element, existedMap: bookmarkMap);
                }

                document.WordBookmarks = bookmarkMap
                    .Where(pair => pair.Value.Valid)
                    .ToDictionary(pair => pair.Key, pair => pair.Value);
            }
        }

        public static void WordTables(Document document)
        {
            if (!document.HasTables)
            {
                document.WordTables = document.RootElement.Elements()
                    .SelectMany(element => Table.SearchFrom(element))
                    .ToList();
            }
        }
    }
}
