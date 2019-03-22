using DocumentFormat.OpenXml.Packaging;
using MSWordLite.Orders;
using MSWordLite.Processes;
using MSWordLite.Tasks;
using System;
using System.IO;
using System.Threading.Tasks;

namespace MSWordLite
{
    public class WordProcess
    {
        public static async Task<IGenerateTask> Run(IGenerateTask task)
        {
            await _run(task);
            return task;
        }

        private static async Task _run(IGenerateTask task)
        {
            if (!task.Valid)
            {
                throw new Exception("both templatePath and template property are empty");
            }

            await _retrieveTemplate(task);
            await _openAndProcess(task);
            await _writeOutputFile(task);
        }

        private static async Task _retrieveTemplate(IGenerateTask task)
        {
            if ((task.Template == null || task.Template.Length == 0) && !string.IsNullOrEmpty(task.TemplatePath))
            {
                if (!File.Exists(task.TemplatePath))
                {
                    throw new FileNotFoundException("template file not found", task.TemplatePath);
                }

                await Task.Run(() => task.Template = File.ReadAllBytes(task.TemplatePath));
            }
        }

        private static async Task _openAndProcess(IGenerateTask task)
        {
            using (var newDocumentStream = new MemoryStream())
            {
                await Task.Run(() =>
                {
                    newDocumentStream.Write(task.Template, 0, task.Template.Length);
                    using (var wordDocument = WordprocessingDocument.Open(newDocumentStream, isEditable: true))
                    {
                        var document = new Document(wordDocument);
                        _initializeAndProcessOrder(task, document);
                    }
                    _generateOutput(task, newDocumentStream);
                });
            }
        }

        private static async Task _writeOutputFile(IGenerateTask task)
        {
            if (!string.IsNullOrEmpty(task.OutputPath) && task.Output != null && task.Output.Length > 0)
            {
                await Task.Run(() => File.WriteAllBytes(task.OutputPath, task.Output));
            }
        }

        private static void _initializeAndProcessOrder(IGenerateTask task, Document document)
        {
            if (!_initializeOrders(task, document))
            {
                task.State = ProcessState.Failure;
                return;
            }

            task.State = ProcessState.Initilized;

            if (!_processingOrders(task, document))
            {
                task.State = ProcessState.Failure;
                return;
            }

            task.State = ProcessState.Success;
        }

        private static bool _initializeOrders(IGenerateTask task, Document document)
        {
            foreach (IOrder order in task.Orders)
            {
                var result = OrderFactory.Initialize(order, document);
                if (!result.Success)
                {
                    task.Error = result.Error;
                    return false;
                }
            }
            return true;
        }

        private static bool _processingOrders(IGenerateTask task, Document document)
        {
            foreach (IOrder order in task.Orders)
            {
                var result = OrderFactory.Process(order, document);
                if (!result.Success)
                {
                    task.Error = result.Error;
                    return false;
                }
            }
            return true;
        }

        private static void _generateOutput(IGenerateTask task, MemoryStream newDocumentStream)
        {
            if (task.State == ProcessState.Success)
            {
                task.Output = newDocumentStream.ToArray();
            }
        }
    }
}
