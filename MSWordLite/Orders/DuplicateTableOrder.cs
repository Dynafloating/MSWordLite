using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace MSWordLite.Orders
{
    public class DuplicateTableOrder : IOrder
    {
        /// <summary>
        /// Index of template table
        /// </summary>
        public int TableId { get; set; }

        /// <summary>
        /// Contents to replace
        /// </summary>
        public List<Dictionary<string, string>> ReplaceContents { get; set; } = new List<Dictionary<string, string>>();

        /// <summary>
        /// Check if this order is valid.
        /// </summary>
        public bool Valid => ReplaceContents != null && TableId >= 0;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableId">Index of template table.</param>
        /// <param name="replaceContents">Contents to replace.</param>
        public DuplicateTableOrder(int tableId, List<Dictionary<string, string>> replaceContents)
        {
            TableId = tableId;
            ReplaceContents = replaceContents;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableId">Index of template table.</param>
        /// <param name="replaceContents">Contents to replace.</param>
        public DuplicateTableOrder(int tableId, List<object> replaceContents)
        {
            TableId = tableId;
            ReplaceContents = replaceContents.Select(o => _convertFromObject(o)).ToList();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableId">Index of template table.</param>
        /// <param name="replaceContents">Contents to replace.</param>
        /// <returns></returns>
        public static IOrder CreateFrom(int tableId, List<Dictionary<string, string>> replaceContents)
        {
            return new DuplicateTableOrder(tableId, replaceContents);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tableId">Index of template table.</param>
        /// <param name="replaceContents">Contents to replace.</param>
        /// <returns></returns>
        public static IOrder CreateFrom(int tableId, List<object> replaceContents)
        {
            return new DuplicateTableOrder(tableId, replaceContents);
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
