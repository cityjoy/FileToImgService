using System;
using System.Collections.Generic; 
using System.Windows.Forms;
using O2S.Components.PDFRender4NET;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Drawing.Drawing2D;
namespace FileToImgService
{
    public class PdfHandle
    {
        public enum Definition
        {
            One = 1, Two = 2, Three = 3, Four = 4, Five = 5, Six = 6, Seven = 7, Eight = 8, Nine = 9, Ten = 10
        }

        /// <summary>
        /// 将PDF文档转换为图片的方法
        /// </summary>
        /// <param name="pdfInputPath">PDF文件路径</param>
        /// <param name="imageOutputPath">图片输出路径</param>
        /// <param name="imageName">生成图片的名字</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换</param>
        /// <param name="imageFormat">设置所需图片格式</param>
        /// <param name="definition">设置图片的清晰度，数字越大越清晰</param>

        public void ConvertPDF2Image(string pdfInputPath, string imageOutputPath,

            string imageName,  ImageFormat imageFormat, Definition definition)
        {

            PDFFile pdfFile = PDFFile.Open(pdfInputPath);

            if (!Directory.Exists(imageOutputPath))
            {

                Directory.CreateDirectory(imageOutputPath);

            }
            int startPageNum = 1;
            int endPageNum = pdfFile.PageCount; 

            // start to convert each page

            for (int i = startPageNum; i <= endPageNum; i++)
            {

                Bitmap pageImage = pdfFile.GetPageImage(i - 1, 56 * (int)definition);

                pageImage.Save(imageOutputPath + imageName + i.ToString() + "." + imageFormat.ToString(), imageFormat);

                pageImage.Dispose();

            }

            pdfFile.Dispose();

        }

        //public static void Main(string[] args)
        //{

        //    ConvertPDF2Image("F:\\Events.pdf", "F:\\", "A",  ImageFormat.Jpeg, Definition.One);

        //}

        /// <summary>
        /// 将PDF文档转换为图片的方法
        /// </summary>
        /// <param name="pdfInputPath">PDF文件路径</param>
        /// <param name="imageOutputPath">图片输出路径</param>
        /// <param name="imageName">生成图片的名字</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换</param>
        /// <param name="imageFormat">设置所需图片格式</param>
        /// <param name="definition">设置图片的清晰度，数字越大越清晰</param>

        //public Tuple<bool,string> ConvertPDF2ImageMew(string pdfInputPath, string imageOutputPath,

        //    string imageName, ImageFormat imageFormat, Definition definition, int attachmentId,string sourcePath)
        //{

        //    try
        //    {
        //        PDFFile pdfFile = PDFFile.Open(pdfInputPath);

        //        if (!Directory.Exists(imageOutputPath))
        //        {

        //            Directory.CreateDirectory(imageOutputPath);

        //        }
        //        int startPageNum = 1;
        //        int endPageNum = pdfFile.PageCount;
        //        var fileSavePath = string.Empty;
        //        // start to convert each page
        //        var aaa = DateTime.Now.ToString("yyyy-MMddhhmmxx");
        //        StringBuilder sb = new StringBuilder();
        //        for (int i = startPageNum; i <= endPageNum; i++)
        //        {

        //            Bitmap pageImage = pdfFile.GetPageImage(i - 1, 56 * (int)definition);
        //            fileSavePath = imageOutputPath +"\\"+aaa+ i.ToString() + "." + imageFormat.ToString();
        //            pageImage.Save(fileSavePath, imageFormat);

        //            pageImage.Dispose();
        //            sb.Append(string.Format("insert into PreviewFile(AttachmentId,SavePath,CreateTime)values({0},'{1}','{2}');", attachmentId, fileSavePath.Replace(sourcePath, "").Replace("\\", "/"), DateTime.Now));
        //        }

        //        pdfFile.Dispose();
        //        var result=  SqlHelper.ExecteNonQueryText(sb.ToString(), null);
        //        return Tuple.Create<bool, string>(result > 0, result > 0 ? "" : pdfInputPath + "数据库操作失败");
        //    }
        //    catch (Exception ex)
        //    {
        //        return Tuple.Create<bool, string>(false, pdfInputPath + "执行失败：" + ex.Message.ToString());
        //    }
        //}


        ///// <summary>
        ///// 将PDF文档转换为图片的方法
        ///// </summary>
        ///// <param name="pdfInputPath">PDF文件路径</param>
        ///// <param name="imageOutputPath">图片输出完整路径(包括文件名)</param>
        ///// <param name="startPageNum">从PDF文档的第几页开始转换</param>
        ///// <param name="endPageNum">从PDF文档的第几页开始停止转换</param>
        ///// <param name="imageFormat">设置所需图片格式</param>
        ///// <param name="definition">设置图片的清晰度，数字越大越清晰</param>
        //public Tuple<bool, string> ConvertPdf2Image0921(string pdfInputPath, string imageOutputPath,
        //     ImageFormat imageFormat, int definition, int attachmentId, string sourcePath)
        //{

        //    try
        //    {
        //        PDFFile pdfFile = PDFFile.Open(pdfInputPath);
        //        int startPageNum = 0; int endPageNum = pdfFile.PageCount; 
        //        if (startPageNum <= 0)
        //        {
        //            startPageNum = 1;
        //        }

        //        if (endPageNum > pdfFile.PageCount)
        //        {
        //            endPageNum = pdfFile.PageCount;
        //        }

        //        if (startPageNum > endPageNum)
        //        {
        //            int tempPageNum = startPageNum;
        //            startPageNum = endPageNum;
        //            endPageNum = startPageNum;
        //        }
        //        var fileSavePath = string.Empty;
        //        var bitMap = new Bitmap[endPageNum];
        //        StringBuilder sb = new StringBuilder();
        //        for (int i = startPageNum; i <= endPageNum; i++)
        //        {
        //            Bitmap pageImage = pdfFile.GetPageImage(i - 1, 56 * definition);
        //            Bitmap newPageImage = new Bitmap(pageImage.Width / 4, pageImage.Height / 4);

        //            Graphics g = Graphics.FromImage(newPageImage);
        //            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
        //            //重新画图的时候Y轴减去130，高度也减去130  这样水印就看不到了
        //            g.DrawImage(pageImage, new Rectangle(0, 0, pageImage.Width / 4, pageImage.Height / 4),
        //                new Rectangle(0, 130, pageImage.Width, pageImage.Height - 130), GraphicsUnit.Pixel);

        //            bitMap[i - 1] = newPageImage;
        //            fileSavePath = imageOutputPath + "\\" + i.ToString() + "." + imageFormat.ToString();
        //            pageImage.Save(fileSavePath, imageFormat);
        //            g.Dispose();
        //            sb.Append(string.Format("insert into PreviewFile(AttachmentId,SavePath,CreateTime)values({0},'{1}','{2}');", attachmentId, fileSavePath.Replace(sourcePath, "").Replace("\\", "/"), DateTime.Now));
        //        }

        //        ////合并图片
        //        //var mergerImg = MergerImg(bitMap);
        //        ////保存图片
        //        //mergerImg.Save(imageOutputPath, imageFormat);
        //        pdfFile.Dispose();
        //        var result = SqlHelper.ExecteNonQueryText(sb.ToString(), null);
        //        return Tuple.Create<bool, string>(result > 0, result > 0 ? "" : pdfInputPath + "数据库操作失败");
        //    }
        //    catch (Exception ex)
        //    {
        //        return Tuple.Create<bool, string>(false, pdfInputPath + "执行失败：" + ex.Message.ToString());
        //    }
           
        //}

        /// <summary>
        /// 合并图片
        /// </summary>
        /// <param name="maps"></param>
        /// <returns></returns>
        private static Bitmap MergerImg(params Bitmap[] maps)
        {
            int i = maps.Length;
            
            if (i == 0)
                throw new Exception("图片数不能够为0");
            else if (i == 1)
                return maps[0];

            //创建要显示的图片对象,根据参数的个数设置宽度
            Bitmap backgroudImg = new Bitmap(maps[0].Width, i * maps[0].Height);
            Graphics g = Graphics.FromImage(backgroudImg);
            //清除画布,背景设置为白色
            g.Clear(System.Drawing.Color.White);
            for (int j = 0; j < i; j++)
            {
                g.DrawImage(maps[j], 0, j * maps[j].Height, maps[j].Width, maps[j].Height);
            }
            g.Dispose();
            return backgroudImg;
        }
         
    }

}