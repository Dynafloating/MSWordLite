namespace MSWordLite.Orders
{
    /// <summary>
    /// Replace bookmark by an image file.
    /// </summary>
    public class InsertImageOrder : IOrder
    {
        /// <summary>
        /// Target bookmark key.
        /// </summary>
        public string BookmarkKey { get; set; }

        /// <summary>
        /// Image data in byte array format.
        /// </summary>
        public byte[] ImageContent { get; set; }

        /// <summary>
        /// Image width (px).
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// Image height (px).
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Image content-type (ex. image/png).
        /// </summary>
        public string ContentType { get; set; }

        /// <summary>
        /// Check if this order is valid.
        /// </summary>
        public bool Valid => !string.IsNullOrEmpty(BookmarkKey) && ImageContent != null && ImageContent.Length > 0 &&
            Width > 0 && Height > 0 && !string.IsNullOrEmpty(ContentType);

        /// <summary>
        /// Create an insert order.
        /// </summary>
        /// <param name="bookmarkKey">Target bookmark key.</param>
        /// <param name="imageContent">Image data in byte array format.</param>
        /// <param name="width">Image width (px).</param>
        /// <param name="height">Image height (px).</param>
        /// <param name="contentType">Image content-type.</param>
        public InsertImageOrder(string bookmarkKey, byte[] imageContent, int width, int height, string contentType)
        {
            BookmarkKey = bookmarkKey;
            ImageContent = imageContent;
            Width = width;
            Height = height;
            ContentType = contentType;
        }

        /// <summary>
        /// Create an order from content.
        /// </summary>
        /// <param name="bookmarkKey">Target bookmark key.</param>
        /// <param name="imageContent">Image data in byte array format.</param>
        /// <param name="width">Image width (px).</param>
        /// <param name="height">Image height (px).</param>
        /// <param name="contentType">Image content-type.</param>
        public static IOrder CreateFrom(string bookmarkKey, byte[] imageContent, int width, int height, string contentType)
        {
            return new InsertImageOrder(bookmarkKey, imageContent, width, height, contentType);
        }
    }
}
