/*************************************************************************************
 * CLR版本：       4.0.30319.42000
 * 类 名 称：      Program
 * 机器名称：      9GX1UOWROPIAEJ4
 * 命名空间：      OfficeToPDF
 * 文 件 名：      Program
 * 创建时间：      2020/12/03 21:51:27
 * 作    者：      Richard Liu
 * 说   明：。。。。。
 * 修改时间：      2020/12/03 21:51:27
 * 修 改 人：      Richard Liu
*************************************************************************************/

using System;
using System.IO;


namespace OfficeToPDF
{
    class Program
    {
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Console.WriteLine(e.ExceptionObject.ToString());
            Environment.Exit(-1); //有此句则不弹“xxx已停止工作”异常对话框
        }


        static void Main(string[] args)
        {
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);

            // 显示软件版本信息
            Version();

            if (args == null)
            {
                Console.WriteLine("文件路径错误");
                Environment.Exit(5);
                return;
            }

            // 输入文件路径
            string inputFile = null;
            inputFile = args[1];

            // 输出文件路径
            string outFile = null;
            if (args.Length == 3)
            {
                outFile = args[2];
            }

            // 获取文件扩展名
            string kind = Path.GetExtension(inputFile).ToLower();
            if (kind == ".doc" || kind == ".docx")
            {
                DocToPDF(inputFile, outFile);
            }
            else if (kind == ".xls" || kind == ".xlsx")
            {
                ExcelToPDF(inputFile, outFile);
            }
            else if (kind == ".ppt" || kind == ".pptx")
            {
                PPTToPDF(inputFile, outFile);
            }
            else if (kind == ".text" || kind == ".txt" || kind == ".rtf")
            {
                DocToPDF(inputFile, outFile);
            }
            else if (kind == ".html" || kind == ".htm" || kind == ".mhtml")
            {
                DocToPDF(inputFile, outFile);
            }
            else
            {
                Console.WriteLine("不支持的文件格式");
                Environment.Exit(5);
            } 
        }

        // DOC转PDF
        private static void DocToPDF(string wpsFilename, string pdfFilename = null)
        {
            // 转换
            int exitCode = 0;
            DocToPDF docToPdf = null;

            try
            {
                docToPdf = new DocToPDF();
                docToPdf.ToPdf(wpsFilename, pdfFilename);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                exitCode = 13;
                Console.WriteLine(@"Did not convert");
            }
            finally
            {
                // 不管转换是否成功都退出WPS
                if (docToPdf != null) { docToPdf.Dispose(); }
            }

            if (exitCode != 0) Environment.Exit(exitCode);
        }

        // EXCEL转PDF
        private static void ExcelToPDF(string wpsFilename, string pdfFilename = null)
        {
            // 转换
            int exitCode = 0;
            ExcelToPDF xlsToPdf = null;
            try
            {
                xlsToPdf = new ExcelToPDF();
                xlsToPdf.ToPdf(wpsFilename, pdfFilename);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                exitCode = 13;
                Console.WriteLine(@"Did not convert");
            }
            finally
            {
                // 不管转换是否成功都退出WPS
                if (xlsToPdf != null) { xlsToPdf.Dispose(); }
            }

            if (exitCode != 0) Environment.Exit(exitCode);
        }

        // PPT转PDF
        private static void PPTToPDF(string wpsFilename, string pdfFilename = null)
        {
            // 转换
            int exitCode = 0;
            PptToPDF pptToPdf = null;
            try
            {
                pptToPdf = new PptToPDF();
                pptToPdf.ToPdf(wpsFilename, pdfFilename);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                exitCode = 13;
                Console.WriteLine(@"Did not convert");
            }
            finally
            {
                // 不管转换是否成功都退出WPS
                if (pptToPdf != null) { pptToPdf.Dispose(); }
            }

            if (exitCode != 0) Environment.Exit(exitCode);

        }

        static void Version()
        {
            Console.WriteLine(@"OfficeToPDF - Convert documents to PDF");
            Console.WriteLine(@"Copyright (c) 2021 Xiamen iLeadTek Technology Co., Ltd");
            Console.WriteLine(@"Copyright (c) 2021 Richard Liu");
            Console.WriteLine(@"Version：1.0." + DateTime.Now.ToString("yyyyMMdd"));
            Console.WriteLine(@"Supported formats：doc、docx、xls、xlsx、ppt、pptx、txt、text、rtf、html、htm、mhtml");
        }
    }
}
