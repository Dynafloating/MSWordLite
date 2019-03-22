using MSWordLite.Orders;
using System.Linq;

namespace MSWordLite.Processes
{
    class ClearBookmarkProcess : OrderProcess<ClearBookmarkOrder>
    {
        public override OrderResult Initialize(Document document)
        {
            return new OrderResult(success: true);
        }

        public override OrderResult Process(Document document)
        {
            if (document.HasBookmarks)
            {
                if (Order.Names.Count > 0)
                {
                    foreach (var bookmark in document.WordBookmarks
                        .Where(p => Order.Names.Contains(p.Key)))
                    {
                        bookmark.Value.Remove();
                    }
                }

                if (Order.Regex != null)
                {
                    foreach (var bookmark in document.WordBookmarks
                        .Where(p => p.Value != null && Order.Regex.IsMatch(p.Key)))
                    {
                        bookmark.Value.Remove();
                    }
                }
            }
            return new OrderResult(success: true);
        }
    }
}
