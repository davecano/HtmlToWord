using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using DAL;
using ICSharpCode.SharpZipLib.Zip;
using MSWord = Microsoft.Office.Interop.Word;
public partial class clipboard : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
       
    }
    public void HtmlToWordByUrl(string url,string title)
    {
        WebBrowser WB = new WebBrowser();//新建内置浏览
        WB.Navigate(url);//加载页面
                         //加载完成
        while (WB.ReadyState != WebBrowserReadyState.Complete)
        {
            System.Windows.Forms.Application.DoEvents();
        }
        //对加载完成的页面进行全选和复制操作
        HtmlDocument doc = WB.Document;
        doc.ExecCommand("SelectAll", false, "");//全选
        doc.ExecCommand("Copy", false, "");//复制
                                           //放入剪切板
        
      //IDataObject iData = Clipboard.GetDataObject();
        SaveWord(title);//保存为word文档
                   //读取文档，下载文档
        Clipboard.Clear();



        //FileStream fs = new FileStream(Server.MapPath("~/UploadFile/test.doc"), FileMode.Open);
        //byte[] bytes = new byte[(int)fs.Length];
        //fs.Read(bytes, 0, bytes.Length);
        //fs.Close();
        //Response.ContentType = "application/octet-stream";
        ////通知浏览器下载文件而不是打开 
        //Response.AddHeader("Content-Disposition", "attachment; filename=htmlfile.doc");
        //Response.BinaryWrite(bytes);
        //WB.Dispose();
        //Response.Flush();
        //Response.End();

    }

    public void SaveWord(string title)
    {
        //string wordstr = wdstr;                   //声明word文档内容
        MSWord.Application wordApp;       //声明word应用程序变量
        MSWord.Document worddoc;          //声明word文档变量    

        //初始化变量
        object Nothing = Missing.Value;                       //COM调用时用于占位
        object format = MSWord.WdSaveFormat.wdFormatDocument; //Word文档的保存格式
        wordApp = new MSWord.ApplicationClass();              //声明一个wordAPP对象
        worddoc = wordApp.Documents.Add(ref Nothing, ref Nothing,
            ref Nothing, ref Nothing);

        //页面设置
        worddoc.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;//设置纸张样式
        worddoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;//排列方式为垂直方向


        //向文档中写入内容（直接粘贴）
        worddoc.Paragraphs.Last.Range.Paste();

        //保存文档
        object path = Server.MapPath("~/UploadFile/"+title+".doc");
        worddoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing,
            ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
            ref Nothing, ref Nothing, ref Nothing, ref Nothing);

        //关闭文档
        worddoc.Close(ref Nothing, ref Nothing, ref Nothing);  //关闭worddoc文档对象
        wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);   //关闭wordApp组对象

    }

    protected void OnClick(object sender, EventArgs e)
    {
        var htmlpath = @Server.MapPath("CreatHtml/");
       var sh=new SQLHelper();
        var dt = sh.ExecuteQuery("select top 2 * from VNews where NewsType='政策法规' order by PublishDate ",
            CommandType.Text);
        for (var i = 0; i < dt.Rows.Count; i++)
        {
            var title = dt.Rows[i]["NewsTitle"].ToString();
            var content = dt.Rows[i]["NewsContent"].ToString();
           CreateHtml(title,content,htmlpath);
           HtmlToWordByUrl(htmlpath+ title + ".html",title);
        }
        PackFiles(@Server.MapPath("/")+"News.zip",@Server.MapPath("UploadFile"));

       Response.Write("success");
    }

  

    private void CreateHtml(string title,string content,string htmlpath)
    {
        //var savepath = @"D:\CreatHtml\" + title + ".html";
        //var htmlpath = @Server.MapPath("CreatHtml/") + title + ".html";
        var fs = new FileStream(htmlpath + title + ".html", FileMode.OpenOrCreate, FileAccess.ReadWrite); //可以指定盘符，也可以指定任意文件名，还可以为word等文件
        var sb=new StringBuilder();
        sb.Append(@"<html>");
        sb.Append(@"<head> <meta charset='utf-8' />");
        sb.Append(@"</head>");
        sb.Append(@"<body>");
        sb.Append(@"<h2>"+title+@"</h2>");
        sb.Append(content);
        sb.Append(@"</body></html>");
        var sw = new StreamWriter(fs); // 创建写入流
        sw.WriteLine(sb); 
        sw.Close(); //关闭文件
    }
    public static void PackFiles(string filename, string directory)
    {
        try
        {
            FastZip fz = new FastZip();
            fz.CreateEmptyDirectories = true;
            fz.CreateZip(filename, directory, true, "");
            fz = null;
        }
        catch (Exception)
        {
            throw;
        }
    }
}