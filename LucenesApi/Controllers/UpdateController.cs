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
    public class UpdateController : ApiController
    {
        static string FileInsert = ConfigurationManager.AppSettings["FileInsert"].ToString();
        static string FilesGet = ConfigurationManager.AppSettings["FilesGet"].ToString();
        static string SelectFlie = ConfigurationManager.AppSettings["SelectFlie"].ToString();
        static string FliePath = ConfigurationManager.AppSettings["FliePath"].ToString();
        static string HtmlPath = ConfigurationManager.AppSettings["HtmlPath"].ToString();
        
        /// <summary>
        /// 查询
        /// </summary>
        /// <returns></returns>
        [System.Web.Http.HttpGet]
        public bool UpdateLucenes(string body)
        {
            Lucenes();
            return true;
        }


        public static List<BodyDto> Lucenes()
        {
            List<BodyDto> BodyDtos = new List<BodyDto>();
            List<HashtableDto> list = new List<HashtableDto>();
            list = FindFolderName(SelectFlie);
            string body = "";
            Analyzer analyzer = new StandardAnalyzer(Lucene.Net.Util.Version.LUCENE_30);
            //IndexWriter writer = new IndexWriter("IndexDirectory", analyzer, true);
            DirectoryInfo dirInfo = System.IO.Directory.CreateDirectory(FilesGet);
            Lucene.Net.Store.Directory directory = Lucene.Net.Store.FSDirectory.Open(dirInfo);
            IndexWriter writer = new IndexWriter(directory, new PanGuAnalyzer(), true, IndexWriter.MaxFieldLength.LIMITED);

            foreach (var item in list)
            {
                if (item.FileName.Split(new char[1] { '.' }).Last() == "doc")
                {
                    string titles = item.FileName.Split(new char[1] { '.' }).First().Split(new char[1] { '\\' }).Last();
                    Titlebody bodys = GetWordContent(item.FileName);
                    string FileNames = item.FileName;
                    long ID = GetSequenceID();
                    AddDocument(writer, titles, bodys.text.Trim().Replace(" ", ""), FileNames, bodys.Title, FliePath+"WORD.png", ID);
                    BodyDto bodyDto = new BodyDto()
                    {
                        fileContent = bodys.text.Trim(),
                        FileName = item.FileName,
                        heardstr = item.FileName.Split(new char[1] { '.' }).First().Split(new char[1] { '\\' }).Last(),
                        filePath = bodys.Title,
                        img = FliePath + "WORD.png",
                        ID = ID
                    };
                    BodyDtos.Add(bodyDto);
                }
                if (item.FileName.Split(new char[1] { '.' }).Last() == "xlsx")
                {
                    string titles = item.FileName.Split(new char[1] { '.' }).First().Split(new char[1] { '\\' }).Last();
                    Excelbody bodys = Excels(item.FileName);
                    string FileNames = item.FileName;
                    long ID = GetSequenceID();
                    AddDocument(writer, titles, bodys.text.Trim().Replace(" ", ""), FileNames, bodys.Title, FliePath+"ECEL.png", ID);
                    BodyDto bodyDto = new BodyDto()
                    {
                        fileContent = bodys.text.Trim(),
                        FileName = item.FileName,
                        heardstr = item.FileName.Split(new char[1] { '.' }).First().Split(new char[1] { '\\' }).Last(),
                        filePath = bodys.Title,
                        img = FliePath + "ECEL.png",
                        ID = ID
                    };
                    BodyDtos.Add(bodyDto);

                }
            }
            writer.Optimize();
            writer.Dispose();
            return BodyDtos;
        }

        public static bool IsNumber(string str)
        {
            if (str == null || str.Length == 0)    //验证这个参数是否为空
                return false;                           //是，就返回False
            ASCIIEncoding ascii = new ASCIIEncoding();//new ASCIIEncoding 的实例
            byte[] bytestr = ascii.GetBytes(str);         //把string类型的参数保存到数组里

            foreach (byte c in bytestr)                   //遍历这个数组里的内容
            {
                if (c < 48 || c > 57)                          //判断是否为数字
                {
                    return false;                              //不是，就返回False
                }
            }
            return true;                                        //是，就返回True
        }




        private static string Preview(string body, string keyword)

        {

            //创建HTMLFormatter,参数为高亮单词的前后缀 
            string key = "";

            foreach (var item in keyword)
            {
                key += item.ToString() + " ";
            }
            var simpleHTMLFormatter =

                new SimpleHTMLFormatter("<font color=\"red\">", "</font>");

            //创建 Highlighter ，输入HTMLFormatter 和 盘古分词对象Semgent 

            var highlighter =

                new Highlighter(simpleHTMLFormatter,

                                new Segment());

            //设置每个摘要段的字符数 

            highlighter.FragmentSize = 300;

            //获取最匹配的摘要段 

            string bodyPreview = highlighter.GetBestFragment(key, body.Trim().ToString());

            if (bodyPreview == null || bodyPreview == "")

                return body;

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
        public static Excelbody Excels(string xlsxName)
        {
            string text = "";
            Excelbody Excelbody = new Excelbody();
            //Save the excel file to PDF format
            string titles = xlsxName.Split(new char[1] { '.' }).First().Split(new char[1] { '\\' }).Last();
            string url = FileInsert + titles + ".html";
            List<HashtableDto> list = FindFolderName(FileInsert);
            if (list.Where(x => x.FileName == url).Count() <= 0)
            {
                ExcelToPdf(xlsxName, FileInsert + titles + ".html");
                text = excel(xlsxName);
            }
            else
            {
                text = excel(xlsxName);
            }
            if (text != null)
            {
                Excelbody = new Excelbody()
                {
                    text = text.Replace("\n", " "),
                    Title = HtmlPath + titles + ".html"
                    // Title = "http://192.168.3.49:8080/group/" + titles + ".html"

                };
            }
            return Excelbody;
        }


        public static string excel(string name)
        {
            DataTable dt = new DataTable();
            Aspose.Cells.Workbook wk = new Aspose.Cells.Workbook(name);
            Worksheet ws = wk.Worksheets[0];

            dt = ws.Cells.ExportDataTable(0, 0, 9, 9);
            string text = "";
            for (int i = 0; i < wk.Worksheets.Count; i++)     //每一sheet
            {

                Cells cells = wk.Worksheets[i].Cells;
                for (int j = 0; j < cells.MaxDataRow + 1; j++)
                {
                    for (int t = 0; t < cells.MaxDataColumn + 1; t++)
                    {

                        text += cells[j, t].StringValue;
                        //一行行的读取数据，插入数据库的代码也可以在这里写
                    }
                }
            }
            return text;
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
            docs.Save(FileInsert + titles + ".html", Aspose.Words.SaveFormat.Html);
            docs.Save(FileInsert + titles + ".pdf", Aspose.Words.SaveFormat.Pdf);
            text = OnCreated(FileInsert + titles + ".pdf");
            Titlebody Titlebody = new Titlebody()
            {
                text = text.Replace("\n", " "),
                Title = HtmlPath + titles + ".html"
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
