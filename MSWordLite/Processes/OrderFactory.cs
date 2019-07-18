using MSWordLite.Orders;
using System.Collections.Generic;

namespace MSWordLite.Processes
{
    class OrderFactory
    {
        private static Dictionary<IOrder, IProcess> _processes { get; set; } = new Dictionary<IOrder, IProcess>();

        public static OrderResult Initialize(IOrder order, Document document)
        {
            var process = _retrieveProcess(order);
            if (process != null)
            {
                _processes.Add(order, process);
                return process.Initialize(document);
            }
            return new OrderResult(success: false, error: "invalid process");
        }

        public static OrderResult Process(IOrder order, Document document)
        {
            return _processes[order].Process(document);
        }

        private static IProcess _retrieveProcess(IOrder order)
        {
            if (order is ClearBookmarkOrder clearBookmarkOrder)
            {
                return new ClearBookmarkProcess() { Order = clearBookmarkOrder };
            }
            else if (order is ReplaceBookmarkOrder replaceBookmarkOrder)
            {
                return new ReplaceBookmarkProcess() { Order = replaceBookmarkOrder };
            }
            else if (order is ExpandTableOrder expandTableOrder)
            {
                return new ExpandTableProcess() { Order = expandTableOrder };
            }
            else if (order is DuplicateTableOrder duplicateTableOrder)
            {
                return new DuplicateTableProcess() { Order = duplicateTableOrder };
            }
            else if (order is InsertImageOrder insertImageOrder)
            {
                return new InsertImageProcess() { Order = insertImageOrder };
            }
            return null;
        }
    }
}
