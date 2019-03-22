using MSWordLite.Orders;
using MSWordLite.Tasks;
using System;
using System.Collections.Generic;

namespace MSWordLite.Cmd
{
    class ReplaceContent
    {
        public string Bookmark0 { get; set; } = "NewText0";
        public string Bookmark1 { get; set; } = "NewText1";
        public string Bookmark2 { get; set; } = "NewText2";
    }

    class Program
    {
        static void Main(string[] args)
        {
            var task = new GenerateTask()
            {
                //TemplatePath = @"C:\Developing\MSWordLite\template_replaceBookmark.docx",
                TemplatePath = @"C:\Developing\MSWordLite\template_expandTable.docx",
                //TemplatePath = @"C:\Developing\MSWordLite\template_duplicateTable.docx",
                OutputPath = @"C:\Developing\MSWordLite\out.docx"
            };

            try
            {
                task.Orders.Add(ReplaceBookmarkOrder.CreateFrom(new ReplaceContent()));

                //task.Orders.Add(ReplaceBookmarkOrder.CreateFrom(new Dictionary<string, string>()
                //{
                //    { "Bookmark0", "NewText0" },
                //    { "Bookmark1", "NewText1" },
                //    { "Bookmark2", "NewText2" }
                //}));

                task.Orders.Add(ExpandTableOrder.CreateFrom(0, new List<List<string>>()
                {
                    new List<string>() { "1", "Data1", "Description1" },
                    new List<string>() { "2", "Data2", "Description2" },
                    new List<string>() { "3", "Data3", "Description3" },
                }));

                //task.Orders.Add(DuplicateTableOrder.CreateFrom(0, new List<Dictionary<string, string>>()
                //{
                //    new Dictionary<string, string>()
                //    {
                //        { "Index", "1" },
                //        { "All", "2" },
                //        { "Insert0", "NewText0" },
                //        { "Insert1", "NewText1" },
                //        { "Insert2", "NewText2" },
                //        { "Insert3", "NewText3" },
                //        { "Insert4", "NewText4 NewText4 NewText4 NewText4" },
                //        { "Insert5", "NewText5 NewText5 NewText5 NewText5 NewText5 NewText5 NewText5" },
                //    },
                //    new Dictionary<string, string>()
                //    {
                //        { "Index", "2" },
                //        { "All", "2" },
                //        { "Insert0", "NewText0-1" },
                //        { "Insert1", "NewText1-1" },
                //        { "Insert2", "NewText2-1" },
                //        { "Insert3", "NewText3-1" },
                //        { "Insert4", "NewText4-1 NewText4 NewText4 NewText4-1" },
                //        { "Insert5", "NewText5-1 NewText5 NewText5 NewText5 NewText5 NewText5 NewText5-1" },
                //    }
                //}));

                task.Orders.Add(new ClearBookmarkOrder());

                WordProcess.Run(task).Wait();

                if (task.State == ProcessState.Failure)
                {
                    Console.WriteLine(task.Error);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            Console.ReadLine();
            task.Dispose();
        }
    }
}
