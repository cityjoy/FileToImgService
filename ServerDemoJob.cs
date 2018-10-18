///***********************************
///*******服务执行的任务**************
///***********************************
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Drawing.Imaging;
using FileToImgService.common;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;
using System.Runtime.InteropServices;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace FileToImgService
{
    public class ServerDemoJob
    {
        int lastCheckId = 0;
        string sourcePath = System.Configuration.ConfigurationSettings.AppSettings.Get("sourcePath");//源文件路劲
        string sourceFilePath = string.Empty;
        string tempPath = System.Configuration.ConfigurationSettings.AppSettings.Get("tempPath");//临时文件路劲
        string tempFilePath = string.Empty;
        string savePath = System.Configuration.ConfigurationSettings.AppSettings.Get("savePath");//图片保存的路劲 
        public void DoJob()
        {
            LogControl.LogInfo("执行服务任务：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            try
            {
                lastCheckId = GetLastId();//最后一次执行的id 
                var filePath = string.Empty;//源文档路劲
                var saveFilePath = string.Empty;//文件保存的路劲
                var searchPath = string.Empty;//文件搜索路劲

                var sqlGetData = string.Format(@"select AttachmentId,SavePath,FileExt from Attachment
                                    where TargetType in(1,2,3) and FileExt in('doc','docx','pdf','ppt','pptx')
                                    and AttachmentId>{0} order by AttachmentId", lastCheckId);
                var data = SqlHelper.ExecuteDataSet(System.Data.CommandType.Text, sqlGetData, null);
                if (data != null && data.Tables.Count > 0)
                {
                    var dataTable = data.Tables[0];
                    if (dataTable != null && dataTable.Rows.Count > 0)
                    {
                        for (var i = 0; i < dataTable.Rows.Count; i++)
                        {
                            sourcePath = System.Configuration.ConfigurationSettings.AppSettings.Get("sourcePath");//源文件路劲
                            tempPath = System.Configuration.ConfigurationSettings.AppSettings.Get("tempPath");//临时文件路劲
                            savePath = System.Configuration.ConfigurationSettings.AppSettings.Get("savePath");//图片保存的路劲
                            lastCheckId = Convert.ToInt32(dataTable.Rows[i]["AttachmentId"]);
                            var pathArr = dataTable.Rows[i]["SavePath"].ToString().Split('/');
                            searchPath = sourcePath + "\\" + pathArr[0] + "\\" + pathArr[1] + "\\OfficeFileToPDF";
                            savePath = sourcePath + "\\" + pathArr[0] + "\\" + pathArr[1] + "\\FileToImg";
                            var ext = dataTable.Rows[i]["FileExt"].ToString().ToUpper();
                            if (savePath.Contains("FileToImg") && Directory.Exists(savePath))
                            {
                                Directory.Delete(savePath, true);
                            }

                            Directory.CreateDirectory(savePath);

                            if (ext == "PDF" || ext == "PPT" || ext == "PPTX")
                            {
                                ext = "PDF";
                                GetPdfImgFromClipboard(searchPath, ext);
                                //PdfToImg(searchPath, ext);
                            }
                            else
                            {
                                WordToImg(searchPath, ext);
                            }
                            ChangeLastId(lastCheckId.ToString());
                        }
                    }
                }
                DoErroeData();
            }
            catch (Exception ex)
            {
                LogControl.LogInfo("执行服务异常：" + ex.Message.ToString());
            }

            LogControl.LogInfo("**********************************执行服务任结束务：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "******************************");
        }

        #region 获取最后一次执行的id
        /// <summary>
        /// 获取最后一次执行的id
        /// </summary>
        /// <returns></returns>
        public int GetLastId()
        {
            try
            {
                var print = "";
                int id = 0;
                if (File.Exists(System.Configuration.ConfigurationManager.AppSettings.Get("lastIdFile") + "\\" + "LactCheckId.txt"))
                {
                    StringBuilder sb = new StringBuilder();
                    using (StreamReader sr = new StreamReader(System.Configuration.ConfigurationManager.AppSettings.Get("lastIdFile") + "\\" + "LactCheckId.txt", Encoding.Default))
                    {
                        while ((print = sr.ReadLine()) != null)
                        {
                            //sb.Append(print);
                            id = Convert.ToInt32(print);
                        }
                    }

                    return id;
                }
                else
                {
                    CheckPath(System.Configuration.ConfigurationManager.AppSettings.Get("lastIdFile"));
                    FileStream fs = new FileStream(System.Configuration.ConfigurationManager.AppSettings.Get("lastIdFile") + "\\" + "LactCheckId.txt", FileMode.OpenOrCreate);
                    StreamWriter sw = new StreamWriter(fs);
                    sw.WriteLine("0");
                    sw.Close();
                    return 0;
                }
            }
            catch (Exception ex)
            {
                return 0;
            }
        }
        #endregion

        #region 记录最后一次执行的id
        /// <summary>
        /// 记录最后一次执行的id
        /// </summary>
        /// <param name="id">最后一笔id</param>
        /// <returns></returns>
        public bool ChangeLastId(string id)
        {
            try
            {
                if (File.Exists(System.Configuration.ConfigurationManager.AppSettings.Get("lastIdFile") + "\\" + "LactCheckId.txt"))
                {
                    using (StreamWriter writer = new StreamWriter(System.Configuration.ConfigurationManager.AppSettings.Get("lastIdFile") + "\\" + "LactCheckId.txt"))
                    {
                        //1,写入文本
                        writer.Write(id);
                    }
                }
                else
                {
                    CheckPath(System.Configuration.ConfigurationManager.AppSettings.Get("lastIdFile"));
                    FileStream fs1 = new FileStream(System.Configuration.ConfigurationManager.AppSettings.Get("lastIdFile") + "\\" + "LactCheckId.txt", FileMode.Create, FileAccess.Write);//创建写入文件 
                    StreamWriter sw = new StreamWriter(fs1);
                    sw.WriteLine(id);//开始写入值
                    sw.Close();
                    fs1.Close();
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        #endregion

        #region 判断文件夹是否存在，不存在则创建
        /// <summary>
        /// 判断文件夹是否存在，不存在则创建
        /// </summary>
        /// <param name="path"></param>
        public void CheckPath(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }
        #endregion

        #region PPT文档转图片
        /// <summary>
        /// PPT文档转图片
        /// </summary>
        /// <param name="pptFileName"></param>
        /// <param name="outPutFilePath"></param>
        public void PPTToImg(string pptFileName, string outPutFilePath)
        {

            Presentation ppt = new Presentation(pptFileName);
            Stream st = new MemoryStream();
            ppt.Save(st, SaveFormat.Pdf);
            Aspose.Pdf.Document document = new Aspose.Pdf.Document(st);
            var device = new Aspose.Pdf.Devices.JpegDevice();

            //默认质量为100，设置质量的好坏与处理速度不成正比，甚至是设置的质量越低反而花的时间越长，怀疑处理过程是先生成高质量的再压缩
            device = new Aspose.Pdf.Devices.JpegDevice(100);
            //遍历每一页转为jpg
            for (var i = 1; i <= document.Pages.Count; i++)
            {
                string savepath = outPutFilePath + "\\" + string.Format("{0}.jpg", i);
                FileStream fs = new FileStream(savepath, FileMode.OpenOrCreate);
                try
                {
                    device.Process(document.Pages[i], fs);
                    fs.Close();
                }
                catch (Exception ex)
                {
                    fs.Close();
                    LogControl.LogInfo(lastCheckId.ToString() + "PPT文档转图片执行失败：" + ex.Message.ToString());

                }
            }
        }

        /// <summary>
        /// PPT文档转PDF
        /// </summary>
        /// <param name="sourcePdf"></param>
        /// <param name="outputPdf"></param>
        /// <returns></returns>
        public bool PPTToPDF(string sourcePdf, string outputPdf)
        {

            try
            {
                new Presentation(sourcePdf).Save(outputPdf, SaveFormat.Pdf);
                return File.Exists(outputPdf);
            }
            catch (Exception)
            {
                return false;
            }
        }




        #endregion

        #region word文档操作
        /// <summary>
        /// word文档操作
        /// </summary>
        /// <param name="searchPath"></param>
        /// <param name="fileExt"></param>
        /// <param name="fileSavePath"></param>
        public void WordToImg(string searchPath, string fileExt)
        {
            try
            {
                if (Directory.Exists(searchPath))
                {
                    DirectoryInfo dir = new DirectoryInfo(searchPath);
                    FileInfo[] dirInfo = dir.GetFiles();
                    if (dirInfo != null && dirInfo.Length > 0)
                    {
                        foreach (var info in dirInfo)
                        {
                            if (info.Extension.ToUpper() == "." + fileExt.ToUpper())
                            {
                                sourceFilePath = info.FullName;
                                CheckPath(tempPath);
                                tempFilePath = tempPath + "\\" + info.Name;
                                if (File.Exists(tempFilePath))
                                {
                                    File.Delete(tempFilePath);
                                }
                                File.Copy(sourceFilePath, tempFilePath);
                                break;
                            }
                        }

                        Word2ImageConverter wordHandle = new Word2ImageConverter();
                        string message = string.Empty; ;
                        var result = wordHandle.ConvertToImageNew(sourcePath, tempFilePath, savePath, lastCheckId, out message);
                        if (!result)
                        {
                            //执行失败不删除临时文件  可以单独处理
                            LogControl.LogInfo(lastCheckId.ToString() + "执行失败：" + message);
                            RecodeErroeData(lastCheckId);
                        }
                        else
                        {
                            File.Delete(tempFilePath);//删除临时文件
                        }
                    }
                }
                else
                {
                    LogControl.LogInfo(lastCheckId.ToString() + "执行失败：找不到路劲：" + searchPath);
                    RecodeErroeData(lastCheckId);
                }
            }
            catch (Exception ex)
            {
                LogControl.LogInfo(lastCheckId.ToString() + "执行失败：" + ex.Message.ToString());
                RecodeErroeData(lastCheckId);
            }
        }
        #endregion

        #region  pdf转img
        /// <summary>
        /// pdf转img
        /// </summary>
        public void PdfToImg(string searchPath, string fileExt)
        {
            if (Directory.Exists(searchPath))
            {
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(searchPath);
                    FileInfo[] dirInfo = dir.GetFiles();
                    string fileName = string.Empty;
                    if (dirInfo != null && dirInfo.Length > 0)
                    {
                        foreach (var info in dirInfo)
                        {
                            if (info.Extension.ToUpper() == "." + fileExt.ToUpper())
                            {
                                sourceFilePath = info.FullName;
                                fileName = info.Name;
                                CheckPath(tempPath);
                                tempFilePath = tempPath + "\\" + info.Name;
                                if (File.Exists(tempFilePath))
                                {
                                    File.Delete(tempFilePath);
                                }
                                File.Copy(sourceFilePath, tempFilePath);
                                break;
                            }

                        }
                        PdfLibNetToImg pdfHandle = new PdfLibNetToImg();

                        var result = pdfHandle.PdfToImg(lastCheckId, sourcePath, tempFilePath, savePath, ImageFormat.Png, 200, 30);
                        //var result = pdfHandle.ConvertPDF2ImageMew(tempFilePath, savePath, fileName.Split('.')[0], ImageFormat.Png, FileToImgService.PdfHandle.Definition.Five, lastCheckId, sourcePath);
                        if (!result)
                        {
                            RecodeErroeData(lastCheckId);
                        }
                        else
                        {
                            File.Delete(tempFilePath);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogControl.LogInfo(lastCheckId.ToString() + "执行失败：" + ex.Message.ToString());
                    RecodeErroeData(lastCheckId);
                }
            }
            else
            {
                LogControl.LogInfo(lastCheckId.ToString() + "执行失败找不到路劲：" + searchPath);
                RecodeErroeData(lastCheckId);
            }
        }
        #endregion]

        #region 从剪切板获取PDF图片
        public void GetPdfImgFromClipboard(string searchPath, string fileExt)
        {
            if (Directory.Exists(searchPath))
            {
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(searchPath);
                    FileInfo[] dirInfo = dir.GetFiles();
                    string fileName = string.Empty;
                    if (dirInfo != null && dirInfo.Length > 0)
                    {
                        foreach (var info in dirInfo)
                        {
                            #region 复制pdf文件到临时目录
                            if (info.Extension.ToUpper() == "." + fileExt.ToUpper())
                            {
                                sourceFilePath = info.FullName;
                                fileName = info.Name;
                                CheckPath(tempPath);
                                tempFilePath = tempPath + "\\" + info.Name;
                                if (File.Exists(tempFilePath))
                                {
                                    File.Delete(tempFilePath);
                                }
                                File.Copy(sourceFilePath, tempFilePath);

                                break;
                            }
                            #endregion

                        }
                        Thread cbThread = new Thread(new ThreadStart(GetImgFromClipboard));
                        cbThread.TrySetApartmentState(ApartmentState.STA);
                        cbThread.Start();
                        cbThread.Join();

                        File.Delete(tempFilePath);
                        
                        //PdfLibNetToImg pdfHandle = new PdfLibNetToImg();

                        //var result = pdfHandle.PdfToImg(lastCheckId, sourcePath, tempFilePath, savePath, ImageFormat.Png, 200, 30);
                        //var result = pdfHandle.ConvertPDF2ImageMew(tempFilePath, savePath, fileName.Split('.')[0], ImageFormat.Png, FileToImgService.PdfHandle.Definition.Five, lastCheckId, sourcePath);
                        //if (!result)
                        //{
                        //    RecodeErroeData(lastCheckId);
                        //}
                        //else
                        //{
                        //    File.Delete(tempFilePath);
                        //}
                    }
                }
                catch (Exception ex)
                {
                    LogControl.LogInfo(lastCheckId.ToString() + "执行失败：" + ex.Message.ToString());
                    RecodeErroeData(lastCheckId);
                }
            }
            else
            {
                LogControl.LogInfo(lastCheckId.ToString() + "执行失败找不到路劲：" + searchPath);
                RecodeErroeData(lastCheckId);
            }
        }

        [STAThread]
        private void GetImgFromClipboard()
        {
            try
            {
                string pdfInputPath = tempFilePath;
                string imageOutputPath = savePath;
                //string pdfInputPath = @"D:\pdfImgs\test.pdf";
                //string imageOutputPath = @"D:\pdfImgs\";
                int startPageNum = 1;
                ImageFormat imageFormat = ImageFormat.Jpeg;
                int resolution = 1;
                Acrobat.CAcroPDDoc pdfDoc = null;
                Acrobat.CAcroPDPage pdfPage = null;
                Acrobat.CAcroRect pdfRect = null;
                Acrobat.CAcroPoint pdfPoint = null;

                // Create the document (Can only create the AcroExch.PDDoc object using late-binding)
                // Note using VisualBasic helper functions, have to add reference to DLL
                pdfDoc = (Acrobat.CAcroPDDoc)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.PDDoc", "");

                // validate parameter
                if (!pdfDoc.Open(pdfInputPath)) { throw new FileNotFoundException(); }
                if (!Directory.Exists(imageOutputPath)) 
                { Directory.CreateDirectory(imageOutputPath); }
                int endPageNum = pdfDoc.GetNumPages();
                if (startPageNum <= 0)
                { startPageNum = 1; }
                if (endPageNum > pdfDoc.GetNumPages() || endPageNum <= 0)
                { endPageNum = pdfDoc.GetNumPages(); }
                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum;
                    startPageNum = endPageNum; endPageNum = startPageNum;
                }
                if (imageFormat == null) 
                { imageFormat = ImageFormat.Jpeg; }
                if (resolution <= 0)
                { resolution = 1; }

                // start to convert each page
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    pdfPage = (Acrobat.CAcroPDPage)pdfDoc.AcquirePage(i - 1);
                    pdfPoint = (Acrobat.CAcroPoint)pdfPage.GetSize();
                    pdfRect = (Acrobat.CAcroRect)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.Rect", "");

                    int imgWidth = (int)((double)pdfPoint.x * resolution);
                    int imgHeight = (int)((double)pdfPoint.y * resolution);

                    pdfRect.Left = 0;
                    pdfRect.right = (short)imgWidth;
                    pdfRect.Top = 0;
                    pdfRect.bottom = (short)imgHeight;

                    // Render to clipboard, scaled by 100 percent (ie. original size)
                    // Even though we want a smaller image, better for us to scale in .NET
                    // than Acrobat as it would greek out small text
                    bool b = pdfPage.CopyToClipboard(pdfRect, 0, 0, (short)(100 * resolution));

                    IDataObject clipboardData = Clipboard.GetDataObject();

                    if (clipboardData.GetDataPresent(DataFormats.Bitmap))
                    {
                        Bitmap pdfBitmap = (Bitmap)clipboardData.GetData(DataFormats.Bitmap);
                        string imtName = i.ToString() + ".jpg";
                        pdfBitmap.Save(imageOutputPath + "\\" + imtName, imageFormat);
                        pdfBitmap.Dispose();
                    }
                }
                if (endPageNum > 0)
                {
                    StringBuilder sb = new StringBuilder();

                    sb.Append(string.Format("delete from PreviewFile where AttachmentId={0};", lastCheckId));
                    for (int num = 1; num <= endPageNum; num++)
                    {
                        string filename = imageOutputPath + "\\" + num.ToString() + ".jpg";
                        sb.Append(string.Format("insert into PreviewFile(AttachmentId,SavePath,CreateTime)values({0},'{1}','{2}');", lastCheckId, filename.Replace(sourcePath, "").Replace("\\", "/"), DateTime.Now));
                        Thread.Sleep(1000);
                    }
                    var result = SqlHelper.ExecteNonQueryText(sb.ToString(), null);
                    if (result > 0)
                    {

                    }
                    else
                    {
                        RecodeErroeData(lastCheckId);
                    }
                }

                pdfDoc.Close();
                Marshal.ReleaseComObject(pdfPage);
                Marshal.ReleaseComObject(pdfRect);
                Marshal.ReleaseComObject(pdfDoc);
                Marshal.ReleaseComObject(pdfPoint);
            }
            catch (Exception ex)
            {
                LogControl.LogInfo("从剪切板获取PDF图片异常lastCheckId:" + lastCheckId +" " + ex.ToString());
                RecodeErroeData(lastCheckId);

            }
        }

        /// <summary>
        /// 将PDF文档转换为图片的方法，你可以像这样调用该方法：ConvertPDF2Image("F:\\A.pdf", "F:\\", "A", 0, 0, null, 0);
        /// 因为大多数的参数都有默认值，startPageNum默认值为1，endPageNum默认值为总页数，
        /// imageFormat默认值为ImageFormat.Jpeg，resolution默认值为1
        /// </summary>
        /// <param name="pdfInputPath">PDF文件路径</param>
        /// <param name="imageOutputPath">图片输出路径</param>
        /// <param name="imageName">图片的名字，不需要带扩展名</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换，默认值为1</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换，默认值为PDF总页数</param>
        /// <param name="imageFormat">设置所需图片格式</param>
        /// <param name="resolution">设置图片的分辨率，数字越大越清晰，默认值为1</param>
        public static void ConvertPDF2Image(string pdfInputPath, string imageOutputPath,
        string imageName, int startPageNum, int endPageNum, ImageFormat imageFormat, double resolution)
        {
            Acrobat.CAcroPDDoc pdfDoc = null;
            Acrobat.CAcroPDPage pdfPage = null;
            Acrobat.CAcroRect pdfRect = null;
            Acrobat.CAcroPoint pdfPoint = null;

            // Create the document (Can only create the AcroExch.PDDoc object using late-binding)
            // Note using VisualBasic helper functions, have to add reference to DLL
            pdfDoc = (Acrobat.CAcroPDDoc)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.PDDoc", "");

            // validate parameter
            if (!pdfDoc.Open(pdfInputPath)) { throw new FileNotFoundException(); }
            if (!Directory.Exists(imageOutputPath)) { Directory.CreateDirectory(imageOutputPath); }
            if (startPageNum <= 0) { startPageNum = 1; } if (endPageNum > pdfDoc.GetNumPages() || endPageNum <= 0) { endPageNum = pdfDoc.GetNumPages(); } if (startPageNum > endPageNum) { int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum; }
            if (imageFormat == null) { imageFormat = ImageFormat.Jpeg; }
            if (resolution <= 0) { resolution = 1; }

            // start to convert each page
            for (int i = startPageNum; i <= endPageNum; i++)
            {
                pdfPage = (Acrobat.CAcroPDPage)pdfDoc.AcquirePage(i - 1);
                pdfPoint = (Acrobat.CAcroPoint)pdfPage.GetSize();
                pdfRect = (Acrobat.CAcroRect)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.Rect", "");

                int imgWidth = (int)((double)pdfPoint.x * resolution);
                int imgHeight = (int)((double)pdfPoint.y * resolution);

                pdfRect.Left = 0;
                pdfRect.right = (short)imgWidth;
                pdfRect.Top = 0;
                pdfRect.bottom = (short)imgHeight;

                // Render to clipboard, scaled by 100 percent (ie. original size)
                // Even though we want a smaller image, better for us to scale in .NET
                // than Acrobat as it would greek out small text
                pdfPage.CopyToClipboard(pdfRect, 0, 0, (short)(100 * resolution));

                IDataObject clipboardData = Clipboard.GetDataObject();

                if (clipboardData.GetDataPresent(DataFormats.Bitmap))
                {
                    Bitmap pdfBitmap = (Bitmap)clipboardData.GetData(DataFormats.Bitmap);
                    pdfBitmap.Save(Path.Combine(imageOutputPath, imageName) + ".jpg", imageFormat);
                    pdfBitmap.Dispose();
                }
            }

            pdfDoc.Close();
            Marshal.ReleaseComObject(pdfPage);
            Marshal.ReleaseComObject(pdfRect);
            Marshal.ReleaseComObject(pdfDoc);
            Marshal.ReleaseComObject(pdfPoint);
        }
        #endregion

        #region 执行执行失败的文件
        /// <summary>
        /// 再执行执行失败的文件
        /// </summary>
        public void CheckErrorFile()
        {
            //DirectoryInfo dir = new DirectoryInfo(tempPath);
            //FileInfo[] dirInfo = dir.GetFiles();
            //string fileName = string.Empty;
            //if (dirInfo != null && dirInfo.Length > 0)
            //{
            //    foreach (var info in dirInfo)
            //    {
            //        if (info.Extension.ToUpper() == ".pdf")
            //        {
            //            sourceFilePath = info.FullName;
            //            fileName = info.Name;
            //            PdfHandle pdfHandle = new PdfHandle();
            //            var result = pdfHandle.ConvertPDF2ImageMew(sourceFilePath, savePath, fileName.Split('.')[0], ImageFormat.Png, FileToImgService.PdfHandle.Definition.Five, lastCheckId, sourcePath);
            //            if (result.Item1)
            //            {
            //                File.Delete(sourceFilePath);
            //            }
            //        }
            //        else
            //        {
            //            Word2ImageConverter wordHandle = new Word2ImageConverter();
            //            var result = wordHandle.ConvertToImageNew(sourcePath, tempPath, savePath, lastCheckId);
            //            if (result.Item1)
            //            {
            //                File.Delete(sourceFilePath);
            //            }
            //        }
            //    }
            //}

        }
        #endregion

        /// <summary>
        /// 执行报错记录的数据
        /// </summary>
        public void DoErroeData()
        {
            LogControl.LogInfo("开始处理异常数据：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            var filePath = string.Empty;//源文档路劲
            var saveFilePath = string.Empty;//文件保存的路劲
            var searchPath = string.Empty;//文件搜索路劲
            var sql = @"select a.Id,a.AttachmentId,b.SavePath,b.FileExt from PreviewFileErrorTable as a
                        left join Attachment as b
                        on b.AttachmentId=a.attachmentid";
            var data = SqlHelper.ExecuteDataSet(System.Data.CommandType.Text, sql, null);
            if (data != null && data.Tables.Count > 0)
            {
                var dataTable = data.Tables[0];
            }
            try
            {
                if (data != null && data.Tables.Count > 0)
                {
                    var dataTable = data.Tables[0];
                    if (dataTable != null && dataTable.Rows.Count > 0)
                    {
                        for (var i = 0; i < dataTable.Rows.Count; i++)
                        {
                            sourcePath = System.Configuration.ConfigurationSettings.AppSettings.Get("sourcePath");//源文件路劲
                            tempPath = System.Configuration.ConfigurationSettings.AppSettings.Get("tempPath");//临时文件路劲
                            savePath = System.Configuration.ConfigurationSettings.AppSettings.Get("savePath");//图片保存的路劲
                            lastCheckId = Convert.ToInt32(dataTable.Rows[i]["AttachmentId"]);
                            var pathArr = dataTable.Rows[i]["SavePath"].ToString().Split('/');
                            searchPath = sourcePath + "\\" + pathArr[0] + "\\" + pathArr[1] + "\\OfficeFileToPDF";
                            savePath = sourcePath + "\\" + pathArr[0] + "\\" + pathArr[1] + "\\FileToImg";
                            var ext = dataTable.Rows[i]["FileExt"].ToString().ToUpper();
                            if (ext == "PDF")
                            {
                                PdfToImg(searchPath, ext);
                            }
                            else
                            {
                                WordToImg(searchPath, ext);
                            }
                            DeleteErrorData(Convert.ToInt32(dataTable.Rows[i]["Id"]));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogControl.LogInfo("处理异常数据报错：" + ex.Message.ToString());
            }
            LogControl.LogInfo("处理异常数据结束：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        }
        /// <summary>
        /// 删除处理完成的异常数据
        /// </summary>
        /// <param name="id"></param>
        public void DeleteErrorData(int id)
        {
            var sql = "delete from PreviewFileErrorTable where id=" + id;
            SqlHelper.ExecteNonQuery(System.Data.CommandType.Text, sql, null);
        }
        /// <summary>
        /// 记录错误文档
        /// </summary>
        /// <param name="attachmentId"></param>
        public void RecodeErroeData(int attachmentId)
        {
            var sql = string.Format("insert into PreviewFileErrorTable(Attachmentid) values({0})", attachmentId);
            SqlHelper.ExecteNonQuery(System.Data.CommandType.Text, sql, null);
        }


        public void test()
        {
            try
            {
                GetCaptureImage();
            }
            catch (Exception ex)
            {
                LogControl.LogInfo("发生异常:" + ex.Message.ToString());
            }

        }
        private static void GetCaptureImage()
        {
            IDataObject iData = Clipboard.GetDataObject();
            Image img = null;
            if (iData != null)
            {
                if (iData.GetDataPresent(DataFormats.Bitmap))
                {
                    img = (Image)iData.GetData(DataFormats.Bitmap);
                }
                else if (iData.GetDataPresent(DataFormats.Dib))
                {
                    img = (Image)iData.GetData(DataFormats.Dib);
                }
            }
            LogControl.LogInfo(img == null ? "Null" : "img");
        }

        /// <summary>       
        /// /// 获取剪切板文本        
        /// </summary>       
        public static string GetClipboard()
        {
            string text = null;
            Thread th = new Thread(new ThreadStart(delegate()
            {
                text = System.Windows.Forms.Clipboard.GetText(); ;
            }));
            th.TrySetApartmentState(ApartmentState.STA);
            th.Start();
            th.Join();
            return text;
        }


    }
}
