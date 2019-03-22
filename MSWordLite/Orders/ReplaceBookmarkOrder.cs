using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace MSWordLite.Orders
{
    /// <summary>
    /// Replace bookmark by text content in document.
    /// </summary>
    public class ReplaceBookmarkOrder : IOrder
    {
        /// <summary>
        /// Content to replace
        /// </summary>
        public IDictionary<string, string> ReplaceContent { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// Check if this order is valid.
        /// </summary>
        public bool Valid => ReplaceContent != null;

        /// <summary>
        /// Create an replace order.
        /// </summary>
        /// <param name="replaceContent">Content to replace</param>
        public ReplaceBookmarkOrder(Dictionary<string, string> replaceContent)
        {
            ReplaceContent = replaceContent;
        }

        /// <summary>
        /// Create an replace order.
        /// </summary>
        /// <param name="replaceContent">Content to replace</param>
        public ReplaceBookmarkOrder(IDictionary<string, string> replaceContent)
        {
            ReplaceContent = replaceContent;
        }

        /// <summary>
        /// Create an replace order.
        /// </summary>
        /// <param name="replaceContent">Content to replace</param>
        public ReplaceBookmarkOrder(object replaceContent)
        {
            ReplaceContent = _convertFromObject(replaceContent);
        }

        /// <summary>
        /// Create an order from content.
        /// </summary>
        /// <param name="replaceContent">Content to replace</param>
        public static IOrder CreateFrom(Dictionary<string, string> replaceContent)
        {
            return new ReplaceBookmarkOrder(replaceContent);
        }

        /// <summary>
        /// Create an order from content.
        /// </summary>
        /// <param name="replaceContent">Content to replace</param>
        public static IOrder CreateFrom(IDictionary<string, string> replaceContent)
        {
            return new ReplaceBookmarkOrder(replaceContent);
        }

        /// <summary>
        /// Create an order from content.
        /// </summary>
        /// <param name="replaceContent">Content to replace</param>
        public static IOrder CreateFrom(object replaceContent)
        {
            return new ReplaceBookmarkOrder(replaceContent);
        }

        private static Dictionary<string, string> _convertFromObject(object replaceContent)
        {
            return replaceContent.GetType().GetRuntimeProperties()
                .Where(prop => prop.CanRead)
                .ToDictionary(prop => prop.Name, prop => prop.GetValue(replaceContent, null))
                .Where(pair => pair.Value != null)
                .ToDictionary(pair => pair.Key, pair => System.Convert.ToString(pair.Value));
        }
    }
}
