namespace MSWordLite.Processes
{
    interface IProcess
    {
        OrderResult Initialize(Document document);
        OrderResult Process(Document document);
    }
}
