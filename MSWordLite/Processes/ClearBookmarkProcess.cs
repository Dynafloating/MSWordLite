using MSWordLite.Orders;
using System.Linq;

namespace MSWordLite.Processes
{
    class ClearBookmarkProcess : OrderProcess<ClearBookmarkOrder>
    {
        private bool _filterByNames => Order.Names.Count > 0;
        private bool _filterByRegex => Order.Regex != null;

        public override OrderResult Initialize(Document document)
        {
            return new OrderResult(success: true);
        }

        public override OrderResult Process(Document document)
        {
            if (document.HasBookmarks)
            {
                if (_filterByNames || _filterByRegex)
                {
                    if (_filterByNames)
                    {
                        foreach (var bookmark in document.WordBookmarks
                            .Where(p => Order.Names.Contains(p.Key)))
                        {
                            bookmark.Value.Remove();
                        }
                    }

                    if (_filterByRegex)
                    {
                        foreach (var bookmark in document.WordBookmarks
                            .Where(p => p.Value != null && Order.Regex.IsMatch(p.Key)))
                        {
                            bookmark.Value.Remove();
                        }
                    }
                }
                else
                {
                    foreach (var bookmark in document.WordBookmarks)
                    {
                        bookmark.Value.Remove();
                    }
                }
            }
            return new OrderResult(success: true);
        }
    }
}
