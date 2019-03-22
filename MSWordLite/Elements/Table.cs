using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Linq;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WordTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace MSWordLite.Elements
{
    /// <summary>
    /// Word 文件表格
    /// </summary>
    class Table
    {
        public WordTable WordTable { get; set; }
        public IEnumerable<TableRow> Rows => WordTable.ChildElements
            .Where(child => child is WordTableRow).Select(child => new TableRow(child as WordTableRow));
        public TableRow FirstRow => Rows.FirstOrDefault();
        public TableRow LastRow => Rows.LastOrDefault();
        public bool Valid => WordTable != default(WordTable);

        public Table(WordTable table)
        {
            WordTable = table;
        }

        public Table Append(TableRow row)
        {
            WordTable.InsertAfter(row.WordRow, LastRow.WordRow);
            return this;
        }

        public static List<Table> SearchFrom(OpenXmlElement documentElement)
        {
            return documentElement.Elements()
                .Where(element => element is WordTable)
                .Select(element => new Table(element as WordTable))
                .ToList();
        }
    }
}
