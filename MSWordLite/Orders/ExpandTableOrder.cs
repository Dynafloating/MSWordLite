using System.Collections.Generic;

namespace MSWordLite.Orders
{
    /// <summary>
    /// Table exapnd process. To use, point a table with its id, process will append data after the last row.
    /// </summary>
    public class ExpandTableOrder : IOrder
    {
        /// <summary>
        /// Index of table
        /// </summary>
        public int TableId { get; set; }

        /// <summary>
        /// Datas to be add.
        /// </summary>
        public List<List<string>> Content { get; set; }

        /// <summary>
        /// Is this order was valided.
        /// </summary>
        public bool Valid => TableId >= 0 && Content != null;

        /// <summary>
        /// Initialize an expand order.
        /// </summary>
        /// <param name="tableId">Index of table.</param>
        /// <param name="content">Contents to append.</param>
        public ExpandTableOrder(int tableId, List<List<string>> content)
        {
            TableId = tableId;
            Content = content;
        }

        /// <summary>
        /// Create an order from table's id and, contents to append.
        /// </summary>
        /// <param name="tableId">Index of table.</param>
        /// <param name="content">Contents to append.</param>
        /// <returns></returns>
        public static IOrder CreateFrom(int tableId, List<List<string>> content)
        {
            return new ExpandTableOrder(tableId, content);
        }
    }
}
