using MSWordLite.Orders;
using System.Linq;
using Table = MSWordLite.Elements.Table;

namespace MSWordLite.Processes
{
    class ExpandTableProcess : OrderProcess<ExpandTableOrder>
    {
        private Table _targetTable { get; set; }

        public override OrderResult Initialize(Document document)
        {
            return new OrderResult(success: true);
        }

        public override OrderResult Process(Document document)
        {
            if (document.WordTables.Count() <= Order.TableId)
            {
                return new OrderResult(success: false, error: "invalid tableId");
            }

            _targetTable = document.WordTables.ElementAt(Order.TableId);
            var rowCloneTargetNumber = _targetTable.Rows.Count() - 1;
            foreach (var rowContent in Order.Content)
            {
                _targetTable.Append(_targetTable.Rows.ElementAt(rowCloneTargetNumber).Clone().ReplaceText(rowContent));
            }
            _targetTable.Rows.ElementAt(rowCloneTargetNumber).WordRow.Remove();

            return new OrderResult(success: true);
        }
    }
}
