using MSWordLite.Orders;
using System.Collections.Generic;

namespace MSWordLite.Tasks
{
    /// <summary>
    /// Word generate task interface
    /// </summary>
    public interface IGenerateTask
    {
        /// <summary>
        /// template path with file name and extension
        /// </summary>
        string TemplatePath { get; set; }

        /// <summary>
        /// template file byte array
        /// </summary>
        byte[] Template { get; set; }

        /// <summary>
        /// output path after generated, file name and extension is needed.
        /// </summary>
        string OutputPath { get; set; }

        /// <summary>
        /// output file byte array
        /// </summary>
        byte[] Output { get; set; }

        /// <summary>
        /// error message during process
        /// </summary>
        string Error { get; set; }

        /// <summary>
        /// state of task
        /// </summary>
        ProcessState State { get; set; }

        /// <summary>
        /// orders to process
        /// </summary>
        List<IOrder> Orders { get; set; }

        /// <summary>
        /// if this task is valid
        /// </summary>
        bool Valid { get; }
    }
}
