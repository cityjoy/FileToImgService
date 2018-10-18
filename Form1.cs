 
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing; 
using System.Text; 
using System.Windows.Forms;

namespace FileToImgService
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            new Aspose.Words.License().SetLicense(LicenseHelper.License.LStream);//去除水印

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ServerDemoJob server = new ServerDemoJob();
            server.DoJob();
            ////初始化一个PdfDocument类实例,并加载PDF文档
            //PdfDocument doc = new PdfDocument();
            //doc.LoadFromFile(@"D:\Temp\tempFIle\复旦大学2016年度毕业生就业质量报告.pdf");
            ////遍历PDF每一页
            //for (int i = 0; i < doc.Pages.Count; i++)
            //{
            //    //将PDF页转换成Bitmap图形
            //    System.Drawing.Image bmp = doc.SaveAsImage(i);
            //    //将Bitmap图形保存为Png格式的图片
            //    string fileName = string.Format("Page-{0}.Jpeg", i + 1);
            //    bmp.Save(fileName, System.Drawing.Imaging.ImageFormat.Jpeg);
            //}   
            //new Aspose.Words.License().SetLicense(LicenseHelper.License.LStream);
            //ServerDemoJob job = new ServerDemoJob();
            //job.DoJob();
        }
    }
}
