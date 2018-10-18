using PDFLibNet;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO; 
using System.Text;
using System.Threading; 

namespace FileToImgService.common
{
    public class PdfLibNetToImg
    {
        /// <summary>
        /// pdf分页转图片
        /// </summary>
        /// <param name="attachmentId">附件id</param>
        /// <param name="sourcePath">文件原始路劲</param>
        /// <param name="fileInputPath">文件路劲（临时文件路劲）</param>
        /// <param name="fileSavePath">文件要保存的路劲</param>
        /// <param name="imageFormat">图片格式</param>
        /// <param name="DPI">图片分辨率（数值越大质量越高）</param>
        /// <param name="definition">清晰度（数值越大质量越高）</param>
        /// <returns></returns>
        public bool PdfToImg(int attachmentId, string sourcePath, string fileInputPath, string fileSavePath, ImageFormat imageFormat, int DPI, int definition)
        {
            try
            {
                List<int> list = new List<int>();
                PDFWrapper wrapper = new PDFWrapper();
                wrapper.LoadPDF(fileInputPath);
                if (!Directory.Exists(fileSavePath))
                {
                    Directory.CreateDirectory(fileSavePath);
                } 
                var pageCount = wrapper.PageCount;
                StringBuilder sb = new StringBuilder();
                if (pageCount > 0)
                {
                    sb.Append(string.Format("delete from PreviewFile where AttachmentId={0};", attachmentId));
                    for (int num = 1; num <= pageCount; num++)
                    {
                        string filename = fileSavePath + "\\" + num.ToString() + "." + imageFormat.ToString();
                        if (File.Exists(filename))
                        {
                            File.Delete(filename);
                        }
                        wrapper.ExportJpg(filename, num, num, (double)DPI, definition);
                        sb.Append(string.Format("insert into PreviewFile(AttachmentId,SavePath,CreateTime)values({0},'{1}','{2}');", attachmentId, fileSavePath.Replace(sourcePath, "").Replace("\\", "/"), DateTime.Now));
                        Thread.Sleep(1000);
                    }
                    wrapper.Dispose();
                    var result = SqlHelper.ExecteNonQueryText(sb.ToString(), null);
                    return result > 0;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception exception)
            {
                throw exception; 
            }
        }
    }
}
