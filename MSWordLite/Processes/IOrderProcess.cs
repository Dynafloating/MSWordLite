using MSWordLite.Orders;

namespace MSWordLite.Processes
{
    interface IOrderProcess<TOrder> : IProcess where TOrder : IOrder
    {
        TOrder Order { get; set; }
    }
}
