using MSWordLite.Orders;
using System;
using System.Collections.Generic;
using System.IO;

namespace MSWordLite.Tasks
{
    /// <summary>
    /// Word 文件產出作業
    /// </summary>
    public class GenerateTask : IGenerateTask, IDisposable
    {
        /// <summary>
        /// 文件範本路徑
        /// </summary>
        public string TemplatePath { get; set; }

        /// <summary>
        /// 範本檔案
        /// </summary>
        public byte[] Template { get; set; }

        /// <summary>
        /// 產出路徑
        /// </summary>
        public string OutputPath { get; set; }

        /// <summary>
        /// 產出檔案
        /// </summary>
        public byte[] Output { get; set; }

        /// <summary>
        /// 錯誤訊息
        /// </summary>
        public string Error { get; set; }

        /// <summary>
        /// 執行狀態
        /// </summary>
        public ProcessState State { get; set; }

        /// <summary>
        /// 產出時執行的命令
        /// </summary>
        public List<IOrder> Orders { get; set; } = new List<IOrder>();

        /// <summary>
        /// 此作業是否符合執行的必要條件
        /// </summary>
        public bool Valid => !string.IsNullOrEmpty(TemplatePath) || Template != null && Template.Length > 0;

        public void Dispose()
        {
            if (!string.IsNullOrEmpty(OutputPath) && File.Exists(OutputPath))
            {
                File.Delete(OutputPath);
            }

            State = ProcessState.Disposed;
        } 
    }
}
