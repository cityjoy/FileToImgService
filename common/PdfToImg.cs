using O2S.Components.PDFRender4NET;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO; 
using System.Text; 

namespace FileToImgService.common
{
    public  class PdfToImage
    {
        public  MemoryStream GetPdfImagePageStream(string pdfInputPath, int pageIndex, ImageFormat format, int width = 1600, int height = 2560, int quality = 10)
        {
            try
            {
                //pdf处理插件
                PDFFile pdfFile = PDFFile.Open(pdfInputPath);
                int total = pdfFile.PageCount;

                #region 防止异常参数
                if (pageIndex < 0)
                {
                    pageIndex = 0;
                }
                if (pageIndex > total)
                {
                    pageIndex = total - 1;
                }
                if (quality < 1)
                {
                    quality = 1;
                }
                if (quality > 10)
                {
                    quality = 10;
                }
                if (width <= 0)
                {
                    width = 1;
                }

                if (height <= 0)
                {
                    height = 1;
                }
                #endregion

                //pdf转换图片
                SizeF pageSize = pdfFile.GetPageSize(pageIndex);

                Bitmap pageImage = pdfFile.GetPageImage(pageIndex, 56 * quality);

                MemoryStream ms = new MemoryStream();

                pageImage.Save(ms, format);

                //原图
                Image img = Image.FromStream(ms, true);

                double ratio = (double)width / (double)height;

                double oRatio = (double)img.Width / (double)img.Height;

                int sbWidth = 0;

                int sbHeight = 0;

                int outX = 0;
                int outY = 0;

                if (oRatio < ratio)
                {
                    sbWidth = (int)(img.Width * ((double)height / (double)(img.Height)));
                    sbHeight = height;

                    outX = (width - sbWidth) / 2;
                }
                else
                {
                    sbHeight = (int)(img.Height * ((double)width / (double)(img.Width)));
                    sbWidth = width;

                    outY = (height - sbHeight) / 2;
                }

                //缩放
                Image sbImg = new Bitmap(sbWidth, sbHeight);
                Graphics sbGra = Graphics.FromImage(sbImg);
                sbGra.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                sbGra.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                sbGra.Clear(Color.White);
                sbGra.DrawImage(img, new System.Drawing.Rectangle(0, 0, sbWidth, sbHeight), new System.Drawing.Rectangle(0, 0, img.Width, img.Height), System.Drawing.GraphicsUnit.Pixel);

                //补白
                Image outImg = new System.Drawing.Bitmap(width, height);
                Graphics outGra = System.Drawing.Graphics.FromImage(outImg);
                outGra.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                outGra.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                outGra.Clear(Color.White);
                outGra.DrawImage(sbImg, new System.Drawing.Rectangle(outX, outY, sbWidth, sbHeight), new System.Drawing.Rectangle(0, 0, sbWidth, sbHeight), System.Drawing.GraphicsUnit.Pixel);

                MemoryStream outMs = new MemoryStream();

                outImg.Save(outMs, format);

                sbImg.Dispose();
                outImg.Dispose();
                img.Dispose();

                return outMs;

            }
            catch (Exception ex)
            {

            }

            return new MemoryStream();
        }

        public static MemoryStream GetPdfImagePageStream(Stream stream, int pageIndex, ImageFormat format, int width = 1600, int height = 2560, int quality = 10)
        {
            try
            {
                //pdf处理插件
                PDFFile pdfFile = PDFFile.Open(stream);
                int total = pdfFile.PageCount;

                #region 防止异常参数
                if (pageIndex < 0)
                {
                    pageIndex = 0;
                }
                if (pageIndex > total)
                {
                    pageIndex = total - 1;
                }
                if (quality < 1)
                {
                    quality = 1;
                }
                if (quality > 10)
                {
                    quality = 10;
                }
                if (width <= 0)
                {
                    width = 1;
                }

                if (height <= 0)
                {
                    height = 1;
                }
                #endregion

                //pdf转换图片
                SizeF pageSize = pdfFile.GetPageSize(pageIndex);

                Bitmap pageImage = pdfFile.GetPageImage(pageIndex, 56 * quality);

                MemoryStream ms = new MemoryStream();

                pageImage.Save(ms, format);

                //原图
                Image img = Image.FromStream(ms, true);

                double ratio = (double)width / (double)height;

                double oRatio = (double)img.Width / (double)img.Height;

                int sbWidth = 0;

                int sbHeight = 0;

                int outX = 0;
                int outY = 0;

                if (oRatio < ratio)
                {
                    sbWidth = (int)(img.Width * ((double)height / (double)(img.Height)));
                    sbHeight = height;

                    outX = (width - sbWidth) / 2;
                }
                else
                {
                    sbHeight = (int)(img.Height * ((double)width / (double)(img.Width)));
                    sbWidth = width;

                    outY = (height - sbHeight) / 2;
                }

                //缩放
                Image sbImg = new Bitmap(sbWidth, sbHeight);
                Graphics sbGra = Graphics.FromImage(sbImg);
                sbGra.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                sbGra.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                sbGra.Clear(Color.White);
                sbGra.DrawImage(img, new System.Drawing.Rectangle(0, 0, sbWidth, sbHeight), new System.Drawing.Rectangle(0, 0, img.Width, img.Height), System.Drawing.GraphicsUnit.Pixel);

                //补白
                Image outImg = new System.Drawing.Bitmap(width, height);
                Graphics outGra = System.Drawing.Graphics.FromImage(outImg);
                outGra.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                outGra.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                outGra.Clear(Color.White);
                outGra.DrawImage(sbImg, new System.Drawing.Rectangle(outX, outY, sbWidth, sbHeight), new System.Drawing.Rectangle(0, 0, sbWidth, sbHeight), System.Drawing.GraphicsUnit.Pixel);

                MemoryStream outMs = new MemoryStream();

                outImg.Save(outMs, format);

                sbImg.Dispose();
                outImg.Dispose();
                img.Dispose();

                return outMs;

            }
            catch (Exception ex)
            {

            }

            return new MemoryStream();
        }
    }
}
