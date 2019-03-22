using System.Collections.Generic;
using System.Linq;
using WordTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WordTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace MSWordLite.Elements
{
    /// <summary>
    /// Word 文件表格行
    /// </summary>
    class TableRow
    {
        public WordTableRow WordRow { get; set; }
        public IEnumerable<TableCell> Cells => WordRow.ChildElements
            .Where(child => child is WordTableCell).Select(child => new TableCell(child as WordTableCell));
        public TableCell FirstCell => Cells.FirstOrDefault();
        public TableCell LastCell => Cells.LastOrDefault();

        public TableRow(WordTableRow row)
        {
            WordRow = row;
        }
        public bool Valid => WordRow != default(WordTableRow);

        public TableRow Clone() =>
            new TableRow((WordTableRow)WordRow.Clone());

        public TableRow ReplaceText(IEnumerable<string> texts)
        {
            for (var i = 0; i < texts.Count(); i++)
            {
                var cell = Cells.ElementAtOrDefault(i);
                var text = texts.ElementAtOrDefault(i);
                if (cell != null && text != "#COPY#")
                {
                    cell.ReplaceText(text);
                }
            }
            return this;
        }
    }
}
