using System.Collections.Generic;

namespace MSWordLite.Orders
{
    /// <summary>
    /// 表格擴增作業，指定一個表格，並將資料附加在表格的最後一行以擴增表格
    /// </summary>
    public class ExpandTableOrder : IOrder
    {
        /// <summary>
        /// Index of table
        /// </summary>
        public int TableId { get; set; }

        /// <summary>
        /// 要加入的內容
        /// </summary>
        public List<List<string>> Content { get; set; }

        /// <summary>
        /// 是否不將表格的第一行視為標題
        /// </summary>
        public bool WithoutHeader { get; set; }

        /// <summary>
        /// 此作業是否有效
        /// </summary>
        public bool Valid => TableId >= 0 && Content != null;

        /// <summary>
        /// 初始化一個表格擴增作業。
        /// </summary>
        /// <param name="tableId">Index of table.</param>
        /// <param name="content">Contents to append.</param>
        /// <param name="withoutHeader">Determine if table contains head row.</param>
        public ExpandTableOrder(int tableId, List<List<string>> content, bool withoutHeader = false)
        {
            TableId = tableId;
            Content = content;
            WithoutHeader = withoutHeader;
        }

        /// <summary>
        /// Create an order from table's id and, contents to append.
        /// </summary>
        /// <param name="tableId">Index of table.</param>
        /// <param name="content">Contents to append.</param>
        /// <param name="withoutHeader">Determine if table contains head row.</param>
        /// <returns></returns>
        public static IOrder CreateFrom(int tableId, List<List<string>> content, bool withoutHeader = false)
        {
            return new ExpandTableOrder(tableId, content, withoutHeader);
        }
    }
}
