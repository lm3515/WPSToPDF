/*************************************************************************************
 * CLR版本：       4.0.30319.42000
 * 类 名 称：      ExcelToPDF
 * 机器名称：      9GX1UOWROPIAEJ4
 * 命名空间：      OfficeToPDF
 * 文 件 名：      ExcelToPDF
 * 创建时间：      2020/12/05 11:51:27
 * 作    者：      Richard Liu
 * 说   明：。。。。。
 * 修改时间：      2020/12/05 11:51:27
 * 修 改 人：      Richard Liu
*************************************************************************************/

using System;
using System.IO;
using Excel;


namespace OfficeToPDF
{
    class ExcelToPDF : IDisposable
    {
        dynamic wps;

        public ExcelToPDF()
        {
            Type type = Type.GetTypeFromProgID("KET.Application");
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
            Console.WriteLine(string.Format(@"正在转换 [{0}]
      -> [{1}]", wpsFilename, pdfFilename));

            object missing = Type.Missing;
            dynamic xls = wps.Application.Workbooks.Open(wpsFilename, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            //保存为Pdf
            xls.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pdfFilename, XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
            xls.Close();
        }

        public void Dispose()
        {
            if (wps != null) { wps.Quit(); }
        }
    }
}
