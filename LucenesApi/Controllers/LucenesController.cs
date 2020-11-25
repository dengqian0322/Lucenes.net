using Aspose.Words;
using Aspose.Words.Saving;
using Lucene.Net.Analysis;
using Lucene.Net.Analysis.Standard;
using Lucene.Net.Documents;
using Lucene.Net.Index;
using Lucene.Net.QueryParsers;
using Lucene.Net.Search;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using Excel = Microsoft.Office.Interop.Excel;
using Aspose.Cells;
using iTextSharp.text.pdf;
using System.Drawing;
using Lucene.Net.Store;
using PanGu.HighLight;
using PanGu;
using System.Text;
using System.Configuration;

namespace LucenesApi.Controllers
{
    public class LucenesController : ApiController
    {

        string FilesGet = ConfigurationManager.AppSettings["FilesGet"].ToString();

        /// <summary>
        /// 查询
        /// </summary>
        /// <returns></returns>
        [System.Web.Http.HttpGet]
        public BodyCount Files(string searchContent)
        {
            BodyCount BodyCounts = new BodyCount();
            List<BodyDto> data = new List<BodyDto>();
            List<Article> Article = new List<Article>();
            if (string.IsNullOrEmpty(searchContent))
            {
                return null;
            }
            else
            {
                string body = "";
                if (SearchIndex(searchContent, FilesGet).Count() <= 0)
                {
                    foreach (var item in searchContent)
                    {
                        body += item.ToString() + " ";
                    }
                    Article = SearchIndex(body, FilesGet);
                }
                else
                {
                    Article = SearchIndex(searchContent, FilesGet);
                }

                foreach (var item in Article)
                {
                    BodyDto BodyDto = new BodyDto()
                    {
                        filePath = item.filePath,
                        fileContent = Preview(item.fileContent, searchContent),
                        FileName = item.FileName,
                        heardstr = item.heardstr,
                        img = item.img,
                        ID = item.ID,
                    };
                    data.Add(BodyDto);
                }
                BodyCounts = new BodyCount()
                {
                    state = 200,
                    total = data.Count(),
                    data = data,
                };
                return BodyCounts;
            }
        }


        /// <summary>
        /// 从索引搜索结果
        /// </summary>
        public static List<Article> SearchIndex(string content, string IndexDic)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            try
            {
                FSDirectory directory = FSDirectory.Open(new DirectoryInfo(IndexDic), new NoLockFactory());
                IndexReader reader = IndexReader.Open(directory, true);
                IndexSearcher search = new IndexSearcher(reader);
                string[] fields = { "content" };

                //创建查询
                PerFieldAnalyzerWrapper wrapper = new PerFieldAnalyzerWrapper(new PanGuAnalyzer());
                wrapper.AddAnalyzer("content", new PanGuAnalyzer());
                QueryParser parser = new MultiFieldQueryParser(Lucene.Net.Util.Version.LUCENE_30, fields, wrapper);
                Query query = parser.Parse(content);
                TopScoreDocCollector collector = TopScoreDocCollector.Create(30, true);//10--查询条数

                search.Search(query, collector);
                var hits = collector.TopDocs().ScoreDocs;

                int numTotalHits = collector.TotalHits;
                List<Article> list = new List<Article>();
                for (int i = 0; i < hits.Length; i++)
                {
                    var hit = hits[i];
                    Lucene.Net.Documents.Document doc = search.Doc(hit.Doc);
                    Article model = new Article()
                    {
                        heardstr = doc.Get("title").ToString(),
                        fileContent = doc.Get("content").ToString(),
                        ID = long.Parse(doc.Get("ID").ToString()),
                        FileName = doc.Get("FileName"),
                        filePath = doc.Get("Url"),
                        img = doc.Get("img")
                    };
                    list.Add(model);
                    //list.Add(SetHighlighter(dicKeywords, model));
                }
                return list;
            }
            catch (Exception ex)
            {
                throw;
            }
        }
        public class Article
        {
            public string heardstr { get; set; }
            public string FileName { get; set; }

            public string fileContent { get; set; }
            public string filePath { get; set; }
            public string img { get; set; }

            public long ID { get; set; }
        }
        private static string Preview(string body, string keyword)
        {

            //创建HTMLFormatter,参数为高亮单词的前后缀 

            var simpleHTMLFormatter =

                new SimpleHTMLFormatter("<font color=\"red\">", "</font>");

            //创建 Highlighter ，输入HTMLFormatter 和 盘古分词对象Semgent 

            var highlighter =new Highlighter(simpleHTMLFormatter, new Segment());

            //设置每个摘要段的字符数 

            highlighter.FragmentSize = 130;

            //获取最匹配的摘要段 

            string bodyPreview = highlighter.GetBestFragment(keyword, body);

            if (bodyPreview == null || bodyPreview == "")
            {
                string key = "";

                foreach (var item in keyword)
                {
                    key += item.ToString() + " ";
                }
                bodyPreview = highlighter.GetBestFragment(key, body);
            }
            return bodyPreview;

        }
        /// <summary>
        /// 获取文件夹下所有文件信息
        /// </summary>
        /// <returns></returns>
        public static List<HashtableDto> FindFolderName(string FileName)
        {
            DirectoryInfo theFolder = new DirectoryInfo(FileName);
            DirectoryInfo[] dirInfo = theFolder.GetDirectories();
            List<HashtableDto> list = new List<HashtableDto>();
            foreach (FileInfo file in theFolder.GetFiles())
            {
                HashtableDto ht = new HashtableDto()
                {
                    FileName = file.FullName,
                };
                list.Add(ht);
            }
            return list;
        }
        static void AddDocument(IndexWriter writer, string title, string content, string FileName, string url, string img, long ID)
        {
            Lucene.Net.Documents.Document document = new Lucene.Net.Documents.Document();
            document.Add(new Field("title", title, Field.Store.YES, Field.Index.ANALYZED));
            document.Add(new Field("content", content, Field.Store.YES, Field.Index.ANALYZED));
            document.Add(new Field("FileName", FileName, Field.Store.YES, Field.Index.ANALYZED));
            document.Add(new Field("Url", url, Field.Store.YES, Field.Index.ANALYZED));
            document.Add(new Field("img", img, Field.Store.YES, Field.Index.ANALYZED));
            document.Add(new Field("ID", ID.ToString(), Field.Store.YES, Field.Index.ANALYZED));
            writer.AddDocument(document);
            Lucene.Net.Documents.Document doc = new Lucene.Net.Documents.Document();
        }

        private static string OnCreated(string filepath)
        {
            try
            {
                string pdffilename = filepath;
                PdfReader pdfReader = new PdfReader(pdffilename);
                int numberOfPages = pdfReader.NumberOfPages;
                string text = string.Empty;

                for (int i = 1; i <= numberOfPages; ++i)
                {
                    iTextSharp.text.pdf.parser.ITextExtractionStrategy strategy = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                    text += iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(pdfReader, i, strategy);
                }
                pdfReader.Close();

                return text;
            }
            catch (Exception ex)
            {
                StreamWriter wlog = File.AppendText(System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "\\mylog.log");
                wlog.WriteLine("出错文件：" + "" + "原因：" + ex.ToString());
                wlog.Flush();
                wlog.Close(); return null;
            }

        }

        public class Excelbody
        {
            public string text { get; set; }
            public string Title { get; set; }
            public string img { get; set; }
        }

        public class BodyCount
        {
            public int state { get; set; }
            public int total { get; set; }
            public List<BodyDto> data { get; set; }
        }
        public class BodyDto
        {
            public string heardstr { get; set; }
            public string FileName { get; set; }

            public string fileContent { get; set; }
            public string filePath { get; set; }
            public string img { get; set; }

            public long ID { get; set; }

        }
        public class HashtableDto
        {
            public string FileName { get; set; }
        }

        /// <summary>
        /// 读取 word文档 返回内容
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static Titlebody GetWordContent(string wordFile)
        {
            string text = "";
            Aspose.Words.Document docs = new Aspose.Words.Document(wordFile);
            string titles = wordFile.Split(new char[1] { '.' }).First().Split(new char[1] { '\\' }).Last();
            //保存为PDF文件，此处的SaveFormat支持很多种格式，如图片，epub,rtf 等等
            docs.Save(@"D:\LucenesApi\group\" + titles + ".html", Aspose.Words.SaveFormat.Html);
            docs.Save(@"D:\LucenesApi\group\" + titles + ".pdf", Aspose.Words.SaveFormat.Pdf);
            text = OnCreated(@"D:\LucenesApi\group\" + titles + ".pdf");
            Titlebody Titlebody = new Titlebody()
            {
                text = text.Replace("\n", " "),
                Title = "http://192.168.3.49:8081/group/" + titles + ".html"
                // Title = "http://192.168.3.49:8080/group/" + titles + ".html"
            };

            return Titlebody;

        }

        public class Titlebody
        {
            public string text { get; set; }
            public string Title { get; set; }
            public string img { get; set; }
        }


        public static long GetSequenceID()
        {
            try
            {
                long tick = ToTimestamp(DateTime.Now, 13);
                long id = DateTime.Now.Year + DateTime.Now.Minute;
                if (id < 10)
                {
                    tick = Convert.ToInt64(tick.ToString() + "0" + id.ToString());
                }
                else
                {
                    tick = Convert.ToInt64(tick.ToString() + id.ToString());
                }
                return tick;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static long ToTimestamp(DateTime dt, int length)
        {
            int x = 10000000;
            if (length == 13)
            {
                x = 10000;
            }
            return (dt.ToUniversalTime().Ticks - 626311977000000000) / x;
        }

        /// <summary>
        /// 将excel转换为html
        /// </summary>
        /// <param name="excelFile">".xls", ".xlsx"类型的文件路径</param>
        /// <param name="pdfFilePath">生成的PDF文件</param>
        /// <returns></returns>
        public static string ExcelToPdf(string excelFile, string pdfSavePath)
        {
            bool isPass = false;
            //string pdfSavePath = string.Empty;
            string msg = string.Empty;
            //excel转换为pdf
            Aspose.Cells.Workbook document = new Aspose.Cells.Workbook(excelFile);

            Aspose.Cells.Style style = document.Styles[document.Styles.Add()];
            style.ShrinkToFit = true;

            int cnt = document.Worksheets.Count;
            for (int i = 0; i < cnt; i++)
            {
                Aspose.Cells.Worksheet sheet = document.Worksheets[i];
                sheet.IsPageBreakPreview = true;

                //sheet.AutoFitColumns();
                //sheet.AutoFitRows();

                sheet.PageSetup.FooterMargin = 0;
                sheet.PageSetup.HeaderMargin = 0;

                //2019-10-12 17:55:55   修改，解决excel文件预览表格时撕裂到第二页了
                sheet.PageSetup.RightMargin = 0;
                sheet.PageSetup.LeftMargin = 0;
                sheet.PageSetup.CenterHorizontally = true;
            }
            foreach (Aspose.Cells.Worksheet p in document.Worksheets)
            {
                p.PageSetup.Zoom = 10;
                p.PageSetup.FitToPagesWide = 1;
                p.PageSetup.FitToPagesTall = 0;
            }

            document.Save(pdfSavePath, Aspose.Cells.SaveFormat.Html);
            isPass = true;
            return pdfSavePath;
        }
    }
}
