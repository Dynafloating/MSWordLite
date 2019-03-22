using MSWordLite.Orders;
using System.Linq;
using Table = MSWordLite.Elements.Table;

namespace MSWordLite.Processes
{
    class ExpandTableProcess : OrderProcess<ExpandTableOrder>
    {
        private Table _targetTable { get; set; }
        private int _rowCloneTargetNumber => Order.WithoutHeader ? 0 : 1;

        public override OrderResult Initialize(Document document)
        {
            Initializer.WordTables(document);
            if (document.WordTables.Count <= Order.TableId)
            {
                return new OrderResult(success: false, error: "invalid tableId");
            }

            _targetTable = document.WordTables.ElementAt(Order.TableId);
            return new OrderResult(success: true);
        }

        public override OrderResult Process(Document document)
        {
            foreach (var rowContent in Order.Content)
            {
                _targetTable.Append(_targetTable.Rows.ElementAt(_rowCloneTargetNumber).Clone().ReplaceText(rowContent));
            }
            _targetTable.Rows.ElementAt(_rowCloneTargetNumber).WordRow.Remove();

            return new OrderResult(success: true);
        }
    }
}
