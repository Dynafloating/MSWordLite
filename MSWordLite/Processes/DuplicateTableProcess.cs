using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MSWordLite.Elements;
using MSWordLite.Orders;
using System.Linq;
using Table = MSWordLite.Elements.Table;

namespace MSWordLite.Processes
{
    class DuplicateTableProcess : OrderProcess<DuplicateTableOrder>
    {
        private Table _templateTable { get; set; }

        public override OrderResult Initialize(Document document)
        {
            if (document.WordTables.Count() <= Order.TableId)
            {
                return new OrderResult(success: false, error: "invalid tableId");
            }

            _templateTable = document.WordTables.ElementAt(Order.TableId);
            return new OrderResult(success: true);
        }

        public override OrderResult Process(Document document)
        {
            var parent = _templateTable.WordTable.Parent;
            OpenXmlElement currentTable = null;

            for (var index = 0; index < Order.ReplaceContents.Count; index++)
            {
                var replaceContent = Order.ReplaceContents[index];
                var newTable = _templateTable.WordTable.CloneNode(true);

                var bookmarkMapInNewTable = Bookmark.SearchFrom(newTable);
                foreach (var bookmark in bookmarkMapInNewTable.Select(p => p.Value))
                {
                    if (replaceContent.ContainsKey(bookmark.Start.Name))
                    {
                        bookmark.Replace(replaceContent[bookmark.Start.Name]);
                    }
                    
                    bookmark.Start.Name = $"{bookmark.Start.Name}-dt-{Order.TableId}-{index}";
                    document.WordBookmarks.Add(bookmark.Start.Name, bookmark);
                }

                if (currentTable == null)
                {
                    currentTable = newTable;
                    parent.InsertAfter(newTable, _templateTable.WordTable);
                }
                else
                {
                    parent.InsertAfter(newTable, currentTable);
                    parent.InsertAfter(new Paragraph(), currentTable);
                    currentTable = newTable;
                }
            }

            _templateTable.WordTable.Remove();

            return new OrderResult(success: true);
        }
    }
}
