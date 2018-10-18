using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing.Imaging;
using System.IO;
using System.Drawing;
using OMCS.Engine.WhiteBoard;
using ESBasic;
using Schematrix;

namespace FileToImgService
{
    /*  
     * 
     * 将pdf、ppf、word转换给图片的组件有很多，这里仅使用Aspose组件（试用版）作为示例。
     * 
     * Aspose官网：www.aspose.com， 请支持和购买正版Aspose组件。
     * 
     */

    #region 图片转换器工厂 -> 将被注入到OMCS的多媒体管理器IMultimediaManager的ImageConverterFactory属性
    /// <summary>
    /// 图片转换器工厂。
    /// </summary>
    public class ImageConverterFactory : IImageConverterFactory
    {
        public IImageConverter CreateImageConverter(string extendName)
        {
            if (extendName == ".doc" || extendName == ".docx")
            {
                return new Word2ImageConverter();
            }

            if (extendName == ".pdf")
            {
                return new Pdf2ImageConverter();
            }

            if (extendName == ".ppt" || extendName == ".pptx")
            {
                return new Ppt2ImageConverter();
            }

            if (extendName == ".rar")
            {
                return new Rar2ImageConverter();
            }

            return null;
        }

        public bool Support(string extendName)
        {
            return extendName == ".doc" || extendName == ".docx" || extendName == ".pdf" || extendName == ".ppt" || extendName == ".pptx" || extendName == ".rar";
        }


        public List<string> GetSupportedFileTypes()
        {
            List<string> list = new List<string>();
            list.Add(".doc");
            list.Add(".docx");
            list.Add(".pdf");
            list.Add(".ppt");
            list.Add(".pptx");
            list.Add(".rar");
            return list;
        }
    }
    #endregion

    #region 将word文档转换为图片
    public class Word2ImageConverter : IImageConverter
    {
        private bool cancelled = false;
        public event CbGeneric<int, int> ProgressChanged;
        public event CbGeneric ConvertSucceed;
        public event CbGeneric<string> ConvertFailed;

        public void Cancel()
        {
            if (this.cancelled)
            {
                return;
            }

            this.cancelled = true;
        }

        public void ConvertToImage(string originFilePath, string imageOutputDirPath)
        {
            this.cancelled = false;
            ConvertToImage(originFilePath, imageOutputDirPath, 0, 0, null, 220);
        }

        /// <summary>
        /// 将Word文档转换为图片的方法      
        /// </summary>
        /// <param name="wordInputPath">Word文件路径</param>
        /// <param name="imageOutputDirPath">图片输出路径，如果为空，默认值为Word所在路径</param>      
        /// <param name="startPageNum">从PDF文档的第几页开始转换，如果为0，默认值为1</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换，如果为0，默认值为Word总页数</param>
        /// <param name="imageFormat">设置所需图片格式，如果为null，默认格式为PNG</param>
        /// <param name="resolution">设置图片的像素，数字越大越清晰，如果为0，默认值为128，建议最大值不要超过1024</param>
        private void ConvertToImage(string wordInputPath, string imageOutputDirPath, int startPageNum, int endPageNum, ImageFormat imageFormat, int resolution)
        {
            try
            {
                Aspose.Words.Document doc = new Aspose.Words.Document(wordInputPath);

                if (doc == null)
                {
                    throw new Exception("Word文件无效或者Word文件被加密！");
                }

                if (imageOutputDirPath.Trim().Length == 0)
                {
                    imageOutputDirPath = Path.GetDirectoryName(wordInputPath);
                }

                if (!Directory.Exists(imageOutputDirPath))
                {
                    Directory.CreateDirectory(imageOutputDirPath);
                }

                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }

                if (endPageNum > doc.PageCount || endPageNum <= 0)
                {
                    endPageNum = doc.PageCount;
                }

                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }

                if (imageFormat == null)
                {
                    imageFormat = ImageFormat.Png;
                }

                if (resolution <= 0)
                {
                    resolution = 128;
                }

                string imageName = Path.GetFileNameWithoutExtension(wordInputPath);
                Aspose.Words.Saving.ImageSaveOptions imageSaveOptions = new Aspose.Words.Saving.ImageSaveOptions(Aspose.Words.SaveFormat.Png);
                imageSaveOptions.Resolution = resolution;
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    if (this.cancelled)
                    {
                        break;
                    }

                    MemoryStream stream = new MemoryStream();
                    imageSaveOptions.PageIndex = i - 1;
                    string imgPath = Path.Combine(imageOutputDirPath, i.ToString()) + "." + imageFormat.ToString();
                    doc.Save(stream, imageSaveOptions);
                    Image img = Image.FromStream(stream);
                    Bitmap bm = ESBasic.Helpers.ImageHelper.Zoom(img, 0.6f);
                    bm.Save(imgPath, imageFormat);
                    img.Dispose();
                    stream.Dispose();
                    bm.Dispose();

                    System.Threading.Thread.Sleep(200);
                    if (this.ProgressChanged != null)
                    {
                        this.ProgressChanged(i - 1, endPageNum);
                    }
                }

                if (this.cancelled)
                {
                    return;
                }

                if (this.ConvertSucceed != null)
                {
                    this.ConvertSucceed();
                }
            }
            catch (Exception ex)
            {
                if (this.ConvertFailed != null)
                {
                    this.ConvertFailed(ex.Message);
                }
            }
        }


        public bool ConvertToImageNew(string sourcePath, string originFilePath, string imageOutputDirPath, int attachmentId,out string  message)
        {
            this.cancelled = false;
            string errormessage=string.Empty;
            var result = ConvertToImageNew(sourcePath, originFilePath, imageOutputDirPath, 0, 0, null, 220, attachmentId, out errormessage);
            message = errormessage;
            return result;
        }
        /// <summary>
        /// 将Word文档转换为图片的方法      
        /// </summary>
        /// <param name="wordInputPath">Word文件路径</param>
        /// <param name="imageOutputDirPath">图片输出路径，如果为空，默认值为Word所在路径</param>      
        /// <param name="startPageNum">从PDF文档的第几页开始转换，如果为0，默认值为1</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换，如果为0，默认值为Word总页数</param>
        /// <param name="imageFormat">设置所需图片格式，如果为null，默认格式为PNG</param>
        /// <param name="resolution">设置图片的像素，数字越大越清晰，如果为0，默认值为128，建议最大值不要超过1024</param>
        private bool ConvertToImageNew(string sourcePath, string wordInputPath, string imageOutputDirPath, int startPageNum, int endPageNum, ImageFormat imageFormat, int resolution, int attachmentId,out string message)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                System.Threading.Thread.Sleep(200);
                Aspose.Words.Document doc = new Aspose.Words.Document(wordInputPath);

                if (doc == null)
                {
                    throw new Exception("Word文件无效或者Word文件被加密！");
                }

                if (imageOutputDirPath.Trim().Length == 0)
                {
                    imageOutputDirPath = Path.GetDirectoryName(wordInputPath);
                }

                if (!Directory.Exists(imageOutputDirPath))
                {
                    Directory.CreateDirectory(imageOutputDirPath);
                }

                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }

                if (endPageNum > doc.PageCount || endPageNum <= 0)
                {
                    endPageNum = doc.PageCount;
                }

                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }

                if (imageFormat == null)
                {
                    imageFormat = ImageFormat.Png;
                }

                if (resolution <= 0)
                {
                    resolution = 128;
                }

                string imageName = Path.GetFileNameWithoutExtension(wordInputPath);
                Aspose.Words.Saving.ImageSaveOptions imageSaveOptions = new Aspose.Words.Saving.ImageSaveOptions(Aspose.Words.SaveFormat.Png);
                imageSaveOptions.Resolution = resolution;
                sb.Append(string.Format("delete from PreviewFile where AttachmentId={0}", attachmentId));
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    if (this.cancelled)
                    {
                        break;
                    }

                    MemoryStream stream = new MemoryStream();
                    imageSaveOptions.PageIndex = i - 1;
                    string imgPath = Path.Combine(imageOutputDirPath, i.ToString()) + "." + imageFormat.ToString();
                    doc.Save(stream, imageSaveOptions);
                    Image img = Image.FromStream(stream);
                    Bitmap bm = ESBasic.Helpers.ImageHelper.Zoom(img, 0.6f);
                    bm.Save(imgPath, imageFormat);
                    img.Dispose();
                    stream.Dispose();
                    bm.Dispose();

                    System.Threading.Thread.Sleep(200);
                    if (this.ProgressChanged != null)
                    {
                        this.ProgressChanged(i - 1, endPageNum);
                    }
                    sb.Append(string.Format("insert into PreviewFile(AttachmentId,SavePath,CreateTime)values({0},'{1}','{2}');", attachmentId, imgPath.Replace(sourcePath, "").Replace("\\", "/"), DateTime.Now));
                }

                if (this.cancelled)
                {
                    message="文档操作失败；";
                    return false;
                }

                if (this.ConvertSucceed != null)
                {
                    this.ConvertSucceed();
                    //执行数据库插入操作 
                }
                int result = SqlHelper.ExecteNonQueryText(sb.ToString(), null);
                message=result > 0?"操作成功":"数据库操作失败；";
                return result > 0;
            }
            catch (Exception ex)
            {
                if (this.ConvertFailed != null)
                {
                    this.ConvertFailed(ex.Message);
                }
                message=ex.Message.ToString();
                return false;
            }
        }
    }
    #endregion

    #region 将pdf文档转换为图片
    public class Pdf2ImageConverter : IImageConverter
    {
        private bool cancelled = false;
        public event CbGeneric<int, int> ProgressChanged;
        public event CbGeneric ConvertSucceed;
        public event CbGeneric<string> ConvertFailed;

        public void Cancel()
        {
            if (this.cancelled)
            {
                return;
            }

            this.cancelled = true;
        }

        public void ConvertToImage(string originFilePath, string imageOutputDirPath)
        {
            this.cancelled = false;
            ConvertToImage(originFilePath, imageOutputDirPath, 0, 0, 200);
        }

        /// <summary>
        /// 将pdf文档转换为图片的方法      
        /// </summary>
        /// <param name="originFilePath">pdf文件路径</param>
        /// <param name="imageOutputDirPath">图片输出路径，如果为空，默认值为pdf所在路径</param>       
        /// <param name="startPageNum">从PDF文档的第几页开始转换，如果为0，默认值为1</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换，如果为0，默认值为pdf总页数</param>       
        /// <param name="resolution">设置图片的像素，数字越大越清晰，如果为0，默认值为128，建议最大值不要超过1024</param>
        private void ConvertToImage(string originFilePath, string imageOutputDirPath, int startPageNum, int endPageNum, int resolution)
        {
            try
            {
                Aspose.Pdf.Document doc = new Aspose.Pdf.Document(originFilePath);

                if (doc == null)
                {
                    throw new Exception("pdf文件无效或者pdf文件被加密！");
                }

                if (imageOutputDirPath.Trim().Length == 0)
                {
                    imageOutputDirPath = Path.GetDirectoryName(originFilePath);
                }

                if (!Directory.Exists(imageOutputDirPath))
                {
                    Directory.CreateDirectory(imageOutputDirPath);
                }

                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }

                if (endPageNum > doc.Pages.Count || endPageNum <= 0)
                {
                    endPageNum = doc.Pages.Count;
                }

                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }

                if (resolution <= 0)
                {
                    resolution = 128;
                }

                string imageNamePrefix = Path.GetFileNameWithoutExtension(originFilePath);
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    if (this.cancelled)
                    {
                        break;
                    }

                    MemoryStream stream = new MemoryStream();
                    string imgPath = Path.Combine(imageOutputDirPath, imageNamePrefix) + "_" + i.ToString("000") + ".jpg";
                    Aspose.Pdf.Devices.Resolution reso = new Aspose.Pdf.Devices.Resolution(resolution);
                    Aspose.Pdf.Devices.JpegDevice jpegDevice = new Aspose.Pdf.Devices.JpegDevice(reso, 100);
                    jpegDevice.Process(doc.Pages[i], stream);

                    Image img = Image.FromStream(stream);
                    Bitmap bm = ESBasic.Helpers.ImageHelper.Zoom(img, 0.6f);
                    bm.Save(imgPath, ImageFormat.Jpeg);
                    img.Dispose();
                    stream.Dispose();
                    bm.Dispose();

                    System.Threading.Thread.Sleep(200);
                    if (this.ProgressChanged != null)
                    {
                        this.ProgressChanged(i - 1, endPageNum);
                    }
                }

                if (this.cancelled)
                {
                    return;
                }

                if (this.ConvertSucceed != null)
                {
                    this.ConvertSucceed();
                }
            }
            catch (Exception ex)
            {
                if (this.ConvertFailed != null)
                {
                    this.ConvertFailed(ex.Message);
                }
            }
        }
    }
    #endregion

    #region 将ppt文档转换为图片
    public class Ppt2ImageConverter : IImageConverter
    {
        private Pdf2ImageConverter pdf2ImageConverter;
        public event CbGeneric<int, int> ProgressChanged;
        public event CbGeneric ConvertSucceed;
        public event CbGeneric<string> ConvertFailed;

        public void Cancel()
        {
            if (this.pdf2ImageConverter != null)
            {
                this.pdf2ImageConverter.Cancel();
            }
        }

        public void ConvertToImage(string originFilePath, string imageOutputDirPath)
        {
            ConvertToImage(originFilePath, imageOutputDirPath, 0, 0, 200);
        }

        /// <summary>
        /// 将pdf文档转换为图片的方法      
        /// </summary>
        /// <param name="originFilePath">ppt文件路径</param>
        /// <param name="imageOutputDirPath">图片输出路径，如果为空，默认值为pdf所在路径</param>       
        /// <param name="startPageNum">从PDF文档的第几页开始转换，如果为0，默认值为1</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换，如果为0，默认值为pdf总页数</param>       
        /// <param name="resolution">设置图片的像素，数字越大越清晰，如果为0，默认值为128，建议最大值不要超过1024</param>
        private void ConvertToImage(string originFilePath, string imageOutputDirPath, int startPageNum, int endPageNum, int resolution)
        {
            try
            {
                Aspose.Slides.Presentation doc = new Aspose.Slides.Presentation(originFilePath);

                if (doc == null)
                {
                    throw new Exception("ppt文件无效或者ppt文件被加密！");
                }

                if (imageOutputDirPath.Trim().Length == 0)
                {
                    imageOutputDirPath = Path.GetDirectoryName(originFilePath);
                }

                if (!Directory.Exists(imageOutputDirPath))
                {
                    Directory.CreateDirectory(imageOutputDirPath);
                }

                if (startPageNum <= 0)
                {
                    startPageNum = 1;
                }

                if (endPageNum > doc.Slides.Count || endPageNum <= 0)
                {
                    endPageNum = doc.Slides.Count;
                }

                if (startPageNum > endPageNum)
                {
                    int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum;
                }

                if (resolution <= 0)
                {
                    resolution = 128;
                }

                //先将ppt转换为pdf临时文件
                string tmpPdfPath = originFilePath + ".pdf";
                doc.Save(tmpPdfPath, Aspose.Slides.Export.SaveFormat.Pdf);

                //再将pdf转换为图片
                Pdf2ImageConverter converter = new Pdf2ImageConverter();
                converter.ConvertFailed += new CbGeneric<string>(converter_ConvertFailed);
                converter.ConvertSucceed += new CbGeneric(converter_ConvertSucceed);
                converter.ProgressChanged += new CbGeneric<int, int>(converter_ProgressChanged);
                converter.ConvertToImage(tmpPdfPath, imageOutputDirPath);

                //删除pdf临时文件
                File.Delete(tmpPdfPath);

                if (this.ConvertSucceed != null)
                {
                    this.ConvertSucceed();
                }
            }
            catch (Exception ex)
            {
                if (this.ConvertFailed != null)
                {
                    this.ConvertFailed(ex.Message);
                }
            }

            this.pdf2ImageConverter = null;
        }

        void converter_ProgressChanged(int done, int total)
        {
            if (this.ProgressChanged != null)
            {
                this.ProgressChanged(done, total);
            }
        }

        void converter_ConvertSucceed()
        {
            if (this.ConvertSucceed != null)
            {
                this.ConvertSucceed();
            }
        }

        void converter_ConvertFailed(string msg)
        {
            if (this.ConvertFailed != null)
            {
                this.ConvertFailed(msg);
            }
        }
    }
    #endregion

    #region 将图片压缩包解压。（如果课件本身就是多张图片，那么可以将它们压缩成一个rar，作为一个课件）
    /// <summary>
    /// 将图片压缩包解压。（如果课件本身就是多张图片，那么可以将它们压缩成一个rar，作为一个课件）
    /// </summary>
    public class Rar2ImageConverter : IImageConverter
    {
        private bool cancelled = false;
        public event CbGeneric<string> ConvertFailed;
        public event CbGeneric<int, int> ProgressChanged;
        public event CbGeneric ConvertSucceed;

        public void Cancel()
        {
            this.cancelled = true;
        }


        public void ConvertToImage(string rarPath, string imageOutputDirPath)
        {
            try
            {
                Unrar tmp = new Unrar(rarPath);
                tmp.Open(Unrar.OpenMode.List);
                string[] files = tmp.ListFiles();
                tmp.Close();

                int total = files.Length;
                int done = 0;

                Unrar unrar = new Unrar(rarPath);
                unrar.Open(Unrar.OpenMode.Extract);
                unrar.DestinationPath = imageOutputDirPath;

                while (unrar.ReadHeader() && !cancelled)
                {
                    if (unrar.CurrentFile.IsDirectory)
                    {
                        unrar.Skip();
                    }
                    else
                    {
                        unrar.Extract();
                        ++done;

                        if (this.ProgressChanged != null)
                        {
                            this.ProgressChanged(done, total);
                        }
                    }
                }
                unrar.Close();

                if (this.ConvertSucceed != null)
                {
                    this.ConvertSucceed();
                }

            }
            catch (Exception ex)
            {
                if (this.ConvertFailed != null)
                {
                    this.ConvertFailed(ex.Message);
                }
            }
        }


    }
    #endregion
}
