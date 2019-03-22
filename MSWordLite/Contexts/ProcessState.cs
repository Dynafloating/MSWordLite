namespace MSWordLite
{
    /// <summary>
    /// 執行狀態
    /// </summary>
    public enum ProcessState
    {
        /// <summary>
        /// 等待中
        /// </summary>
        Waiting,

        /// <summary>
        /// 初始化中
        /// </summary>
        Initilized,

        /// <summary>
        /// 成功
        /// </summary>
        Success,

        /// <summary>
        /// 失敗
        /// </summary>
        Failure,

        /// <summary>
        /// 產出的檔案已刪除
        /// </summary>
        Disposed
    }
}
