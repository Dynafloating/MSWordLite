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
                TemplatePath = @"C:\Developing\MSWordLite\template_replaceBookmark.docx",
                //TemplatePath = @"C:\Developing\MSWordLite\template_expandTable.docx",
                //TemplatePath = @"C:\Developing\MSWordLite\template_duplicateTable.docx",
                OutputPath = @"C:\Developing\MSWordLite\out.docx"
            };

            try
            {
                //task.Orders.Add(ReplaceBookmarkOrder.CreateFrom(new ReplaceContent()));

                //task.Orders.Add(ReplaceBookmarkOrder.CreateFrom(new Dictionary<string, string>()
                //{
                //    { "Bookmark0", "NewText0" },
                //    { "Bookmark1", "NewText1" },
                //    { "Bookmark2", "NewText2" }
                //}));

                //task.Orders.Add(ExpandTableOrder.CreateFrom(0, new List<List<string>>()
                //{
                //    new List<string>() { "1", "Data1", "Description1" },
                //    new List<string>() { "2", "Data2", "Description2" },
                //    new List<string>() { "3", "Data3", "Description3" },
                //}));

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

                var base64 = "iVBORw0KGgoAAAANSUhEUgAAAWoAAAClCAMAAABCz0cEAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAMAUExURf" +
                    "///wAAAAgICCkpKff37+/m7yEZIRAQGTo6OhAQCK2trWNjY0pKSqWMlBA6Uq1azkpazkpahBBaGealnHtazube3tbe3rVaUubWc+Z75u" +
                    "Zac61apeYx5ubWMeZaMWMpUmNaMbVaGWMpGRla7xlapebWUuZaUmMIUmMIGa3vWkrv3krvWq2U70qt3kqtWq2lWkqtnEqtGa3vGUrvnE" +
                    "rvGUoZ70oZpa2lGa0Zpa0Z73vvnHvvWnuU73vvGXsZpXsZ73ulGUrO3krOWq2EWkrOnErOGa2EGXuEGYSEhKWllHtapRBjUhAQUoScjM" +
                    "7Oxb3Fxa1a70pa70papTFaGea95ualvXta7+ac5rW9tdbW1tbWnHt7czFjUjEQUub3exAxGZQpUuZateYQtealEOYpEJQpGRmM7xmMax" +
                    "mMrRmMKRkpzhkphITO75QIUuZalOYQlOaEEOYIEJQIGRmMzhmMShmMjBmMCBkIzhkIhITOzpRKUtZa5pxahNYQ5tbWENZaEFJaEJRKGQ" +
                    "hazghahHNrc1JaWq3vrXulpXula63Oa62lzkqM70qMa0qMrUqMKa3OKUopzkophK0phK0pznvOrXvOa3ulznvOKXsphHspzq3FlK3vjH" +
                    "ulSq3OSq2EzkqMzkqMSkqMjEqMCK3OCEoIzkoIhK0IhK0IznvOjHvOSnuEznvOCHsIhHsIzrUpUuZ7teYxtealMeYpMbUpGealc+Ypcx" +
                    "mt7xmtaxmtrRmtKRkp7xkppYTv7xnv7xnvaxnvrRnvKeb3xeaEc+YIcxnO7xnOaxnOrRnOKbUIUuZ7lOYxlOaEMeYIMbUIGealUuYpUh" +
                    "mtzhmtShmtjBmtCBkI7xkIpYTvzhnvzhnvShnvjBnvCOaEUuYIUhnOzhnOShnOjBnOCPfWnJRrUvda5r1ahPcQ5vfWEPdaEHNaEJRrGS" +
                    "lazilahLXv77XF73tSe7XvzikQCO/Wxa2MrXuEpXuEUub3EDExEGtKWjo6Wub3Qt737973/wAQCAAACP/3/wAAANThTVsAAAEAdFJOU/" +
                    "////////////////////////////////////////////////////////////////////////////////////////////////////////" +
                    "////////////////////////////////////////////////////////////////////////////////////////////////////////" +
                    "////////////////////////////////////////////////////////////////////////////////////////////////////////" +
                    "///////////////////////////wBT9wclAAAACXBIWXMAACHVAAAh1QEEnLSdAAASqklEQVR4Xu2d3ZKqOBDHJ0ASBObSPBO5GLiF4s" +
                    "KXkfedmnMxVVNo7b8DKgg6gIg6y692z0F3djRNp7/SCW8LCwsLCwsLCwsLCwsLCwsLCwsLCwsLCwsLC/8XOK8uFu5Jqv08yLcfX8ki77" +
                    "uiA1EwYiU3rp/sq7cXJoZHuSzFLOUKf1vMzby/qtsPHVeakaBFkGdp5IW+K4zc8z8pbOWFOnrYwJIAgpU1ySqduyTtPKne+DOo0MW4bF" +
                    "29nJuEPt311tVLA488DZti67+l2BEpFY22ej0zJGkZRNWrGtz7ZEX2l2StcohZQLHt6o15SXGjRdgtUBWwXfZTvfgDeCRpzwkeo9Uqg0" +
                    "571YsWiS9Z3rAsr4yyIWmY6YDl1Ttz4oSCydCpXrXhGbP+ig1x3hHNwtErlz3C3XOXyUxVL7pQoWT+tR94HTBWlkNteGA9YkAwH+71z3" +
                    "VCy/r+E7KGpRakzpGQDwj2IptJfdl8GPZhwYK0evHKYBw+2cLIZfOLmme78tOvwnXB8teXdSSYa2x0ZO8uxgF3w5NIXarrK/B3/FxH4P" +
                    "1S8EzK0sGnNns3b80IXDHrFV7wLWzIi8satlKULieeX6sRyLGgn8PjGs77tX2jliwovRI0bG5RIxfs/ZkcMd/VoPDZQU4uq2g6FsXMbn" +
                    "GYpjqw177qY22eE3KK1beHKflXXs0F0lRrQNb0g7wxj6sXrwf06jCDI9ua2YCEvS11SRRYLHhERjsF8EvHVA0RyLwGBEp9uczUxR6hIX" +
                    "NfNL6GKzzGWom9mzfYI4880M+FFmT9mjFfVIsA0o4IJPLSn3s1ZvCcrYbeW7OG8ZK1J/7B7KOOdGSL3GcyyDPvLmODRxaDjYEnIOvtC+" +
                    "r13mfbo85C1GF1eYBrWvRju0DfQdihLLLqsj8UHyJCfD17HdVD6cguWgaER5lNY2NBPLUZQfoiR2jnj+lieD17rXemfFqS2ruOCIRjnh" +
                    "PiY2JZJ4J9V5dD2HPfyPrV9Hp7zF9ALKxOL5WGRrGrSslUUPljUKR3wuh18LimlVG49YXbyC2+qsszlE+aPW3uQPZj7M0zsrY/XikQ4a" +
                    "Lwq0uAFOaSmnGdBxMvpa4D5o5eCE8y3HsreKFuHCV2tRggFsV5BHKCT13niXbM/2Wd6wr7OHTZTr5Om1m0sWrCnXcVxmPsNqXkSSCndi" +
                    "D3IxL1GgSyxflqIOQVby3SrbUbvIoJSWQt1oMB6Qr27gTlobfHxuplPGO4qg8XKcx89Wruvs7kn4KwMdzUHVKnvxG+GVzVe2m0rIs6mr" +
                    "NerUSP9o8/RMjs2nBhq+erVyc9mxL+Cpms22qYz/mCvYg9omvtcYTysFpOKHc3n62GqGcM4h+PJ1e18apgxmAvYuJ1175HkMjGLA5uTi" +
                    "r6k7x+C94gUsFqiTl3rTlFTS3d/x9SsaqVmxIxQf7WF+9/JmpH1He/wC3OJ+qQvUz5Yhrsuqg9MbpWPxyIetIP28eR8p65oS+oL3glwp" +
                    "pPq302fLX8GvGnlJK2qT0reX3Hj5arDlHz+2zzz2Rt/WcCaEMgY9P+zklBZn6SridZcxlaRXEce/rDi6efmdupbbWX+/nY1kke67uf9a" +
                    "Lrog4b66pc6VzsGC2YMhl8TN0GMrWtxhd21mrkaS083/2yn+92IllLj/VKnrS6PErhhJi4YT98pmCPduRcXladBlVLF/cZE9XlG/8ojz" +
                    "5hlgxy3xzMsZu25un9ti90Thx/8jaXFo5bu5un/huuqe9DCt9T6zXnsCV0tsGk5kw/1coAdXBl9z0riucnUav86KkSqLFoHOZjesgnDA" +
                    "W5fqoaCN8y1m5YnBQnPBXoVXBszPiHD27aZv6+wX2vXkyBV262fhZIuYbd+2jrDzPv3ikUhag/qktqrBV5o02DmgmmtK5KPtUqDDnGWv" +
                    "9zD7yCiUEDiE4mE/bqMPiy2VM2YoQY5mxCL83d5yqCZBa0a0Bhk0zOprruR3oSNbUbVZcwpQGddedmXqIU52RXYNanFA5myS4Z30g2Oa" +
                    "a12e7f3UqbWwdp9f5HiIOFSnasFmPwstMX4d7G3ZodGt7muJl0CmAdn2rJnDbZDLBpCX56WBlA2QdR87DZ1ssRa55w397W08b5PJj0zt" +
                    "2MpjNCey8tm30iw4pb6+CwdIsZbdVKIM4HmesjlNx8WcytWzOugFEDrkbUpDQcY3X5DKRmFts9PT+Z092wKAG6VYkaxqfRfwO56sQLg0" +
                    "DAahdU1jbpaylSrkLfd4UrhGsHfm7brj84waFe9icKrd8wqzHQnkYNAdTQ/BJmoorclWj8v+8IM/GpDl9zta4qe3CcNqk1jzMhrfJQ3i" +
                    "P1Uzb7gfDxiaw1TwTbWUz0GwXZj4Fz0skOXc5RY+C4Be3EmbwufoZT6gh/CZUWG7kqpCxtjbxwKuElSK2fyVpDrRGGGGX6jYjC8KGnPo" +
                    "eVTSA7XzfzedceKshYptzsRJFBGMfrNAnfQ0/rMKdYqfCHHfOYsdUz1fLp1mMYzfMTeer7uZ+oRlxanms6NJHXhwBnWy9dm/ij/avWFF" +
                    "vnZDnssz2jPNribet9kF7DUdhPpNZ7vWKBZEV9cqot6S8TtpudKvZqa6bxUO8U73Kjihh2Q40RN3YUlGHPGEm6I0XHL0D+1J4J19Cstf" +
                    "/3kcQ2VAgT1juqsEO1kZKdfVA9OliUGNo0E4tSbOf9zgg3Ok4H4T4J2uqsf6Z0/4e5ivP7+2DIgniQ4zGro9VmJt3q3HJhVuS5QpRKRn" +
                    "3wmvc6KJdelNUMByijaUt0TaorLuRUpAH1JuIeIKZ5pkcKZLAe8HjFcT7TEcjhmxNpn/RIBjr5CksR4NVQUcMoG3MZ4qaZNw5QuNFKh5" +
                    "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";

                var array = Convert.FromBase64String(base64);

                task.Orders.Add(InsertImageOrder.CreateFrom("Bookmark0", array, 362, 165, "image/png"));

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
