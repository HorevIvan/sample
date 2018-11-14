using Dapper;
using Subtitles.Data;
using System;
using HI;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Collections.Concurrent;

namespace Subtitles.Webapp.Controllers
{
    public class SearchController : BaseController
    {
        public static ConcurrentDictionary<String, ConcurrentDictionary<String, Int64[]>> SearchCache { get; } = new ConcurrentDictionary<string, ConcurrentDictionary<string, long[]>>();

        // GET https://ru.celluloid.guru/search?p=1&q=Работа
        //
        public ActionResult Show()
        {
            var pageSize = 50;

            var query = Request["q"];

            var langDomain = (string)ViewBag.LangDomain;

            var pageNum = 0;

            if (int.TryParse("0" + Request["p"], out int parsedPage)) pageNum = parsedPage;

            pageNum = pageNum <= 0 ? 1 : pageNum;

            var model = new SearchModel { Query = query, Language = langDomain.ToLower()  };

            query = new string(query.Where(ch => Char.IsLetterOrDigit(ch) || ch == ' ' || ch == '\'' || ch == '`').ToArray());

            query = query.Trim(' ');

            long[] phraseDBIDs;

            if (SearchCache.ContainsKey(langDomain) && SearchCache[langDomain].ContainsKey(query))
            {
                phraseDBIDs = SearchCache[langDomain][query];
            }
            else
            {
                phraseDBIDs = SearchByFullText(query, langDomain);

                if (SearchCache.ContainsKey(langDomain).Not()) SearchCache[langDomain] = new ConcurrentDictionary<string, long[]>();

                SearchCache[langDomain][query] = phraseDBIDs;
            }

            model.Total = phraseDBIDs.Count();
            model.Page = pageNum;
            model.PagesCount = (model.Total / pageSize) + (model.Total % pageSize == 0 ? 0 : 1);

            phraseDBIDs = phraseDBIDs.Skip(pageSize * (pageNum - 1)).Take(pageSize).ToArray();

            using (var db = SubtitlesDataContext.Create())
            {
                var phrases =
                    db.Phrases
                        .Include("Lines")
                        .Include("Frame")
                        .Include("Translation")
                        .Include("Translation.Video")
                        .Include("Translation.Video.Film")
                        .Where(ph => phraseDBIDs.Contains(ph.DBID))
                        .ToArray();

                model.Phrases = phrases.ToArray();

                return View(model);
            }
        }

        private static Int64[] SearchByFullText(String query, String langDomain)
        {
            Int64[] phraseDBIDs;

            var fullTextSearchQuery = $"'{query}'";

            var queryParts = query.Split(' ').ToArray();

            if (queryParts.Length > 1)
            {
                fullTextSearchQuery = "'";

                foreach (var queryPart in queryParts)
                {
                    fullTextSearchQuery += $"\"{queryPart}\"";

                    if (queryPart != queryParts.Last()) fullTextSearchQuery += " and ";
                }

                fullTextSearchQuery += "'";
            }

            var sql = ""
                + " SELECT p.DBID from Frames f"
                + " join Phrases p on p.FrameDBID = f.DBID"
                + " join Lines l on l.PhraseDBID = p.DBID"
                + " join Translations t on t.DBID = p.TranslationDBID"
                + " where contains(l.Text, " + fullTextSearchQuery + ") and t.Domain ='" + langDomain + "'"
                + " order by p.DBID";

            var connectionString = SubtitlesDataContext.GetConnectionString();

            using (var connection = new SqlConnection(connectionString))
            {
                phraseDBIDs = connection.Query<long>(sql).ToArray();
            }

            return phraseDBIDs;
        }
    }

    public class SearchModel
    {
        public Phrase[] Phrases { get; set; }
        public String Query { get; set; }
        public Int32 Total { get; internal set; }
        public Int32 PagesCount { get; internal set; }
        public Int32 Page { get; internal set; }
        public String Language { get; internal set; }
    }
}