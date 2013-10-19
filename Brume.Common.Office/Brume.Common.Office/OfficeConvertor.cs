using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using ApplicationClass = Microsoft.Office.Interop.Word.ApplicationClass;

using ET;
using WPP;
using WPS;
using Document = WPS.Document;
using Presentation = WPP.Presentation;
using WdAlertLevel = Microsoft.Office.Interop.Word.WdAlertLevel;
using WdSaveFormat = WPS.WdSaveFormat;
using _Workbook = ET._Workbook;

namespace Brume.Common.Office
{
    /// <summary>
    ///     Office 文档转换类
    ///     Created By Alien
    /// </summary>
    public class OfficeConvertor
    {
        private readonly IFileHelper _fileHelper;

        public OfficeConvertor()
        {
            _fileHelper = new MongoDbFileHelper();
        }

        /// <summary>
        ///     获取转换后的临时文件地址
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string GetTempDocumentUrl(string id)
        {
            if (IsExsit(id))
            {
                return "temp/" + id + ".xps";
            }
            var path = _fileHelper.GetDocumentPath(id).ToLower();
            var result = "";
            if (path.Contains("doc"))
            {
                result = ConvertWordToXps(path);
            }
            else if (path.Contains("xls"))
            {
                result = ConvertExcelToXps(path);
            }
            else if (path.Contains("ppt"))
            {
                result = ConvertPPtToXps(path);
            }
            else if (path.Contains("wps"))
            {
                result = ConvertWpsToXps(path);
            }
            else if (path.Contains("et"))
            {
                result = ConvertEtToXps(path);
            }
            else if (path.Contains("dps"))
            {
                result = ConvertDpsToXps(path);
            }
            return "temp/" + result;
        }

        #region Kingsoft Convertors

        /// <summary>
        ///     获取金山Wps文件转换结果地址
        /// </summary>
        /// <param name="filename">wps文件地址</param>
        private string ConvertWpsToXps(string filename)
        {
            var app = new WPS.ApplicationClass { Visible = false, DisplayAlerts = WpsAlertLevel.wpsAlertsNone };
            try
            {
                Document doc = app.Documents.Open(filename);
                string docFile = filename.Substring(0, filename.LastIndexOf('.')) + ".doc";
                var type = WdSaveFormat.wdFormatDocument as object;
                doc.SaveAs(docFile, ref type);
                //app.PrintOut(ref missing, ref missing, ref missing, ref tFile);
                doc.Close();
                doc = null;
                app.Quit();
                return ConvertWordToXps(docFile);
            }
            catch (Exception)
            {
                return "";
            }
            finally
            {
                app = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                ClearProcesses("wps");
            }

            //return GetFileName(xpsFile);
        }

        /// <summary>
        ///     获取金山表格转换地址
        /// </summary>
        /// <param name="filename">金山表格地址</param>
        /// <returns>xps文件地址</returns>
        private string ConvertEtToXps(string filename)
        {
            var app = new ET.ApplicationClass { Visible = false, DisplayAlerts = false };
            try
            {
                _Workbook workBook = app.Workbooks.Open(filename);
                string xlsFile = filename.Substring(0, filename.LastIndexOf('.')) + ".xls";
                var tfile = xlsFile as object;
                var type = ETFileFormat.etExcel12 as object;
                workBook.SaveAs(tfile, type);
                workBook.Close();
                workBook = null;
                app.Quit();
                app = null;

                ClearProcesses("et");

                return ConvertExcelToXps(xlsFile);
            }
            catch (Exception)
            {
                return "";
            }
            finally
            {
                if (app != null)
                {
                    app.Quit();
                    app = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        ///     转换WPS演示文档到Xps
        /// </summary>
        /// <param name="filename">演示文档路径</param>
        /// <returns>xps文件路径</returns>
        private string ConvertDpsToXps(string filename)
        {
            var app = new WPP.ApplicationClass();
            try
            {
                Presentation pp = app.Presentations.Open(filename);
                string pptFile = filename.Substring(0, filename.LastIndexOf('.')) + ".ppt";
                pp.SaveAs(pptFile, Type.Missing, WpSaveAsFileType.wpSaveAsPowerPoint7);
                pp.Close();
                pp = null;
                app.Quit();
                app = null;
                return ConvertPPtToXps(pptFile);
            }
            catch
            {
                return "";
            }
            finally
            {
                app = null;
                ClearProcesses("WPP");
            }
        }

        #endregion

        #region MS-Office Convertors

        /// <summary>
        ///     Word转Xps
        /// </summary>
        /// <param name="filename">word文件位置</param>
        /// <returns>xps文件位置</returns>
        private string ConvertWordToXps(string filename)
        {
            var app = new ApplicationClass
            {
                Visible = false,
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };
            var file = filename as object;

            try
            {
                var doc = app.Documents.Open(ref file);
                //app.ActivePrinter = GetXpsPrinter()["name"].ToString();
                string xpsFile = filename.Substring(0, filename.LastIndexOf('.')) + ".xps";
                //var tFile = xpsFile as object;
                //app.PrintOut(ref missing, ref missing, ref missing, ref tFile);
                doc.ExportAsFixedFormat(xpsFile, WdExportFormat.wdExportFormatXPS);
                doc.Close();
                doc = null;
                app.Quit();
                return GetFileName(xpsFile);
            }
            catch (Exception)
            {
                return "";
            }
            finally
            {
                app = null;

                ClearProcesses("WINWORD");

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        ///     Excel转XPS
        /// </summary>
        /// <param name="filename">Excel文件位置</param>
        /// <returns>xps文件位置</returns>
        private string ConvertExcelToXps(string filename)
        {
            var app = new Microsoft.Office.Interop.Excel.ApplicationClass {Visible = false, DisplayAlerts = false};
            try
            {
                Workbook wk = app.Workbooks.Open(filename);
                try
                {
                    //app.ActivePrinter = GetXpsPrinter()["name"].ToString();
                    string xpsFile = filename.Substring(0, filename.LastIndexOf('.')) + ".xps";
                    wk.ExportAsFixedFormat(XlFixedFormatType.xlTypeXPS, xpsFile);
                    wk.Close();
                    wk = null;
                    app.Quit();
                    return GetFileName(xpsFile);
                }
                finally
                {
                    if (wk != null)
                    {
                        wk.Close();
                        wk = null;
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                ClearProcesses("EXCEL");
            }
            return "";
        }

        /// <summary>
        ///     PPT转XPS
        /// </summary>
        /// <param name="filename">PPT文件位置</param>
        /// <returns>xps文件位置</returns>
        private string ConvertPPtToXps(string filename)
        {
            var app = new Microsoft.Office.Interop.PowerPoint.ApplicationClass();
            //app.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            try
            {
                var p = app.Presentations.Open(filename,
                    MsoTriState.msoTrue);
                var xpsFile = filename.Substring(0, filename.LastIndexOf('.')) + ".xps";
                p.SaveAs(xpsFile, PpSaveAsFileType.ppSaveAsXPS);
                p.Close();
                p = null;

                app.Quit();
                return GetFileName(xpsFile);
            }
            catch (Exception)
            {
                return "";
            }
            finally
            {
                ClearProcesses("POWERPNT");

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion

        #region Motheds

        /// <summary>
        ///     获取文件名
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns></returns>
        private static string GetFileName(string path)
        {
            var info = new FileInfo(path);
            return info.Name;
        }

        /// <summary>
        ///     获取文件后缀
        /// </summary>
        /// <param name="filename">文件名</param>
        /// <returns></returns>
        private static string GetExtention(string filename)
        {
            int s = filename.LastIndexOf('.');
            int l = filename.Length - s;
            return filename.Substring(s, l);
        }

        /// <summary>
        ///     清理过期临时文件
        /// </summary>
        /// <param name="dir">临时文件夹</param>
        private void ClearFile(string dir)
        {
            var dirInfo = new DirectoryInfo(dir);
            FileInfo[] fileInfos = dirInfo.GetFiles();
            foreach (
                FileInfo info in
                    from info in fileInfos
                    let exTime = DateTime.Now.AddSeconds(-30)
                    where info.CreationTime < exTime
                    select info)
            {
                try
                {
                    File.Delete(info.FullName);
                }
                catch
                {
                }
            }
        }

        /// <summary>
        ///     清理进程
        /// </summary>
        /// <param name="name">进程名</param>
        private void ClearProcesses(string name)
        {
            var ps = Process.GetProcessesByName(name);
            foreach (var process in ps)
            {
                try
                {
                    var time = process.TotalProcessorTime;
                    Thread.Sleep(5);
                    if (time == process.TotalProcessorTime)
                    {
                        process.Kill();
                    }
                }
                catch
                {
                }
            }
        }

        /// <summary>
        ///     是否已经存在转换文件
        /// </summary>
        /// <param name="id">文件Id</param>
        /// <returns></returns>
        private bool IsExsit(string id)
        {
            var filename = id + ".xps";
            var dir = HttpContext.Current.Request.MapPath("~/temp/");
            return File.Exists(dir + filename);
        }

        #endregion
    }
}