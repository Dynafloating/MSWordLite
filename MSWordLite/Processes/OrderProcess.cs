using MSWordLite.Orders;

namespace MSWordLite.Processes
{
    abstract class OrderProcess<TOrder> : IOrderProcess<TOrder> where TOrder : IOrder
    {
        public TOrder Order { get; set; }
        public abstract OrderResult Initialize(Document document);
        public abstract OrderResult Process(Document document);
    }
}
