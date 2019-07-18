using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MSWordLite.Elements
{
    /// <summary>
    /// Word 文件書籤
    /// </summary>
    class Bookmark
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public BookmarkStart Start { get; set; }
        public BookmarkEnd End { get; set; }
        public bool Valid => !string.IsNullOrEmpty(Name) && Start != null && End != null;

        public Bookmark(BookmarkStart start)
        {
            if (start != null)
            {
                Start = start;
                Name = Start.Name;
                Id = Start.Id;
            }
        }

        public Bookmark AppendEnd(BookmarkEnd end)
        {
            End = end;
            return this;
        }

        public bool Replace(string text)
        {
            if (!Valid) { return false; }

            var adeptedText = !string.IsNullOrEmpty(text) ? text : "";
            var splitedText = adeptedText.Split(new string[] { "<br/>" }, StringSplitOptions.None);
            
            _clearBetweenStartAndEnd(Start);

            var run = _createRunFromTexts(Start, splitedText);
            Start.Parent.InsertAfter(run, Start);

            return true;
        }

        public bool InsertRun(Run run)
        {
            if (!Valid) { return false; }

            _clearBetweenStartAndEnd(Start);
            Start.Parent.InsertAfter(run, Start);

            return true;
        }

        private static void _clearBetweenStartAndEnd(BookmarkStart start)
        {
            var elem = start.NextSibling();
            var paragraph = start.Parent as Paragraph;

            while (elem != null && !(elem is BookmarkEnd))
            {
                var nextElem = elem.NextSibling();
                elem.Remove();
                elem = nextElem;
            }

            if (!(elem is BookmarkEnd))
            {
                var pg = paragraph.NextSibling();
                while (pg != null && !(elem is BookmarkEnd))
                {
                    var bookmarkEnd = _findBookmarkEnd(pg, start.Id);
                    if (bookmarkEnd != null)
                    {
                        elem = bookmarkEnd;

                        var nextElem = bookmarkEnd.NextSibling();
                        paragraph.AppendChild(bookmarkEnd.CloneNode(true));
                        while (nextElem != null)
                        {
                            paragraph.AppendChild(nextElem.CloneNode(true));
                            nextElem = nextElem.NextSibling();
                        }

                        pg.Remove();
                    }
                    else
                    {
                        var nextPg = pg.NextSibling();
                        pg.Remove();
                        pg = nextPg;
                    }
                }
            }
        }

        private static BookmarkEnd _findBookmarkEnd(OpenXmlElement documentElement, string id)
        {
            foreach (var elem in documentElement.ChildElements)
            {
                if (!(elem is BookmarkEnd end) || end.Id != id)
                {
                    var result = _findBookmarkEnd(elem, id);
                    if (result != null)
                    {
                        return result;
                    }
                }
                else
                {
                    return end;
                }
            }

            return null;
        }

        private static Run _createRunFromTexts(BookmarkStart start, IEnumerable<string> text)
        {
            var run = new Run();
            var count = text.Count();
            for (var i = 0; i < count; i++)
            {
                run.Append(new Text(text.ElementAt(i)));
                if (i < count - 1)
                {
                    run.Append(new Break());
                }
            }

            var paragraph = start.Parent as Paragraph;
            var pgp = paragraph.ChildElements.Where(child => child is ParagraphProperties).FirstOrDefault();
            if (pgp != null)
            {
                var pgmrp = pgp.ChildElements.Where(child => child is ParagraphMarkRunProperties).FirstOrDefault();
                if (pgmrp != null)
                {
                    run.RunProperties = new RunProperties(pgmrp.OuterXml);
                }
            }

            return run;
        }

        public void Remove()
        {
            if (Start != null && Start.Parent != null)
            {
                Start.Remove();
            }
        }

        public static Dictionary<string, Bookmark> SearchFrom(
            OpenXmlElement documentElement, Dictionary<string, Bookmark> existedMap = null)
        {
            if (existedMap == null)
            {
                existedMap = new Dictionary<string, Bookmark>();
            }

            foreach (var element in documentElement.Elements())
            {
                if (element is BookmarkStart)
                {
                    var start = new Bookmark(element as BookmarkStart);
                    if (!existedMap.ContainsKey(start.Name))
                    {
                        existedMap.Add(start.Name, start);
                    }
                }

                if (element is BookmarkEnd)
                {
                    var end = element as BookmarkEnd;
                    var name = existedMap
                        .Where(item => item.Value.Id == end.Id)
                        .Select(item => item.Value.Name)
                        .FirstOrDefault();

                    if (!string.IsNullOrEmpty(name))
                    {
                        existedMap[name] = existedMap[name].AppendEnd(end);
                    }
                }

                SearchFrom(element, existedMap);
            }

            return existedMap;
        }
    }
}
