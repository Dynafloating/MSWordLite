using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using WordTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace MSWordLite.Elements
{
    /// <summary>
    /// Word 文件表格欄位
    /// </summary>
    class TableCell
    {
        public WordTableCell WordCell { get; set; }
        public IEnumerable<Paragraph> Paragraphs => WordCell.ChildElements
            .Where(child => child is Paragraph).Select(child => (Paragraph)child);
        public IEnumerable<Run> Runs => Paragraphs
            .SelectMany(pg => pg.ChildElements.Where(pgc => pgc is Run).Select(pgc => (Run)pgc));
        public bool Valid => WordCell != default(WordTableCell);

        public TableCell(WordTableCell cell)
        {
            WordCell = cell;
        }

        public TableCell ReplaceText(string text)
        {
            var splitedText = text.Split(new string[] { "<br/>" }, StringSplitOptions.None);
            var pg = Paragraphs.FirstOrDefault();
            var run = Runs.FirstOrDefault();
            if (run != null)
            {
                var t = run.ChildElements[1] as Text;
                t.Text = text;
            }
            else if (pg != null)
            {
                var newRun = new Run();
                var count = splitedText.Count();
                for (var i = 0; i < count; i++)
                {
                    newRun.Append(new Text(splitedText.ElementAt(i)));
                    if (i < count - 1)
                    {
                        newRun.Append(new Break());
                    }
                }

                var pgp = pg.ChildElements.Where(child => child is ParagraphProperties).FirstOrDefault();
                if (pgp != null)
                {
                    var pgmrp = pgp.ChildElements.Where(child => child is ParagraphMarkRunProperties).FirstOrDefault();
                    if (pgmrp != null)
                    {
                        newRun.RunProperties = new RunProperties(pgmrp.OuterXml);
                    }
                }

                pg.AppendChild(newRun);
            }
            return this;
        }
    }
}
