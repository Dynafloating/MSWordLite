namespace MSWordLite.Processes
{
    public class OrderResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }

        public OrderResult(bool success)
        {
            Success = success;
        }

        public OrderResult(bool success, string error)
        {
            Success = success;
            Error = error;
        }
    }
}
