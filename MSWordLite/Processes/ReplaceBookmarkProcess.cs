using MSWordLite.Orders;

namespace MSWordLite.Processes
{
    class ReplaceBookmarkProcess : OrderProcess<ReplaceBookmarkOrder>
    {
        public override OrderResult Initialize(Document document)
        {
            Initializer.Bookmarks(document);
            return new OrderResult(success: true);
        }

        public override OrderResult Process(Document document)
        {
            foreach (var content in Order.ReplaceContent)
            {
                if (document.WordBookmarks.ContainsKey(content.Key))
                {
                    document.WordBookmarks[content.Key].Replace(content.Value);
                }
            }

            return new OrderResult(success: true);
        }
    }
}
