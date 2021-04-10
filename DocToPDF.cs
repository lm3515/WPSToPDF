/*************************************************************************************
 * CLR版本：       4.0.30319.42000
 * 类 名 称：      DocToPDF
 * 机器名称：      9GX1UOWROPIAEJ4
 * 命名空间：      OfficeToPDF
 * 文 件 名：      DocToPDF
 * 创建时间：      2020/12/05 11:51:27
 * 作    者：      Richard Liu
 * 说   明：。。。。。
 * 修改时间：      2020/12/05 11:51:27
 * 修 改 人：      Richard Liu
*************************************************************************************/


using System;
using System.IO;
using Word;

namespace OfficeToPDF
{
    class DocToPDF : IDisposable
    {
        dynamic wps;

        public DocToPDF()
        {
            Type type = Type.GetTypeFromProgID("KWps.Application");
            wps = Activator.CreateInstance(type);
        }

        public void ToPdf(string wpsFilename, string pdfFilename = null)
        {
            if (wpsFilename == null) { throw new ArgumentNullException("wpsFilename"); }

            if (pdfFilename == null)
            {
                pdfFilename = Path.ChangeExtension(wpsFilename, "pdf");
            }

            //忽略警告提示
            wps.DisplayAlerts = false;

            Console.WriteLine(string.Format(@"正在转换 [{0}] -> [{1}]", wpsFilename, pdfFilename));
            //用WPS打开word不显示界面
            dynamic doc = wps.Documents.Open(wpsFilename, Visible: false);
            doc.ExportAsFixedFormat(pdfFilename, WdExportFormat.wdExportFormatPDF);
            doc.Close();
        }

        public void Dispose()
        {
            if (wps != null) 
            { 
                wps.Quit(); 
            }
        }
    }
}
