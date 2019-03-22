using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MSWordLite.Elements;
using MSWordLite.Orders;
using System.Collections.Generic;
using System.Linq;

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
