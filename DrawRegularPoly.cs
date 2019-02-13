using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;

namespace WindowsServiceDemo
{
    /// <summary>
    /// 雷达图
    /// </summary>
    public class DrawRegularPoly
    {


        /// <summary>
        /// 通用雷达图背景
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="len">边数</param>
        /// <param name="cd">对角线的长度</param>
        /// <param name="rotation">旋转角度</param>
        /// <returns></returns>
        public byte[] GetTYRadarBackgroundImageGraph(int width, int height, int len, int cd, double rotation)
        {
            Bitmap bm = new Bitmap(width, height);
            SolidBrush barBrush1 = new SolidBrush(Color.Gray);
            SolidBrush barBrush2 = new SolidBrush(Color.FromArgb(0, 206, 209));

            Graphics g = Graphics.FromImage(bm);//画板
            g.SmoothingMode = SmoothingMode.AntiAlias;//抗锯齿
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            System.Drawing.Font f = new System.Drawing.Font("微软雅黑", 1);
            Pen p = new Pen(barBrush1, 0.1f);
            Pen p1 = new Pen(barBrush2, 2f);

            double val = rotation;//旋转度，如果为0则不旋转，从y轴为0开始

            try
            {

                //画同心多边形===========================================================================================
                for (int j = 1; j <= 5; j++)
                {
                    System.Drawing.Point[] points1 = new System.Drawing.Point[len + 1];
                    for (int i = 0; i < len; i++)
                    {
                        int x = (int)(width / 2 + cd / 5 * j * Math.Cos((double)2 * i * Math.PI / len + val * Math.PI / 180));
                        int y = (int)(height / 2 - cd / 5 * j * Math.Sin((double)2 * i * Math.PI / len + val * Math.PI / 180));
                        points1[i] = new System.Drawing.Point(x, y);
                    }

                    points1[len] = new System.Drawing.Point((int)(width / 2 + cd / 5 * j * Math.Cos(Math.PI * val / 180)), (int)(height / 2 - cd / 5 * j * Math.Sin(Math.PI * val / 180)));

                    g.DrawLines(p, points1);
                }
                MemoryStream stream = new MemoryStream();

                bm.Save(stream, ImageFormat.Png);

                //System.Drawing.Image png = System.Drawing.Image..GetInstance(bm, ImageFormat.Png);

                return stream.GetBuffer();

            }
            finally
            {
                g.Dispose();
                bm.Dispose();
            }
        }


        /// <summary>
        /// 多元智能测评正八边形雷达图
        /// </summary>
        /// <param name="radius">正多边形外切圆的半径</param>
        /// <param name="sideCount">正多边形的边数</param>
        /// <param name="scoreList">正多边形的嵌套个数</param>
        public void Draw8RegularPoly(double radius, int[] scoreList)
        {
            double[] center = new double[] { 250.0, 200.0 };//正多边形外切圆的圆心固定
            int sideCount = 8;
            if (scoreList.Length != sideCount)
            {
                return;
            }
            // 每条边对应的圆心角角度，精确为浮点数。使用弧度制，360度角为2派
            double arc = 2 * Math.PI / sideCount;
            // 为多边形创建所有的顶点列表
            var pointList = new List<double[]>();
            System.Drawing.Point[] points = new System.Drawing.Point[8];
            Bitmap bm = new Bitmap(500, 400);
            Graphics g = Graphics.FromImage(bm);//画板
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            g.SmoothingMode = SmoothingMode.AntiAlias;  //使绘图质量最高，即消除锯齿
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.CompositingQuality = CompositingQuality.HighQuality;
            System.Drawing.Font f = new System.Drawing.Font("微软雅黑", 1);
            Pen p = new Pen(Color.Gray, 0.1f); //指定颜色和粗细
            //g.DrawLine(p, new System.Drawing.Point(0, 0), new System.Drawing.Point(800, 0));//X轴
            //g.DrawLine(p, new System.Drawing.Point(0, 0), new System.Drawing.Point(0, 800));//Y轴
            for (int j = 0; j < 3; j++)
            {
                radius -= 50;
                for (int i = 0; i < sideCount; i++)
                {
                    var curArc = arc * i; // 当前点对应的圆心角角度
                    double[] pt = new double[3];
                    // 就是简单的三角函数正余弦根据圆心角和半径算点坐标。这里都取整就行
                    pt[0] = center[0] + Math.Round((radius * Math.Cos(curArc)), 2);
                    pt[1] = center[1] + Math.Round((radius * Math.Sin(curArc)), 2);
                    pt[2] = 0;
                    pointList.Add(pt);

                    points[i] = new System.Drawing.Point((int)pt[0], (int)pt[1]);
                }
                SolidBrush barBrush1 = new SolidBrush(Color.Gray);
                SolidBrush barBrush2 = new SolidBrush(Color.FromArgb(0, 206, 209));


                g.DrawPolygon(p, points);//画多边形
                if (j == 0)//画对角线
                {
                    g.DrawLine(p, points[0], points[4]);
                    g.DrawLine(p, points[1], points[5]);
                    g.DrawLine(p, points[2], points[6]);
                    g.DrawLine(p, points[3], points[7]);
                }
            }


            Point point0 = new Point(100 + 15 * (10 - scoreList[0]), 200);
            Point point1 = new Point(100 + 150 - ((int)(Math.Cos(Math.PI / 4) * (15 * scoreList[1]))), 50 + 150 - ((int)(Math.Sin(Math.PI / 4) * (15 * scoreList[1]))));
            Point point2 = new Point(100 + 150, 50 + 15 * (10 - scoreList[2]));
            Point point3 = new Point(100 + 150 + ((int)(Math.Cos(Math.PI / 4) * (15 * scoreList[3]))), 50 + 150 - ((int)(Math.Sin(Math.PI / 4) * (15 * scoreList[3]))));
            Point point4 = new Point(100 + 150 + 15 * scoreList[4], 200);
            Point point5 = new Point(100 + 150 + ((int)(Math.Cos(Math.PI / 4) * (15 * scoreList[5]))), 50 + 150 + ((int)(Math.Sin(Math.PI / 4) * (15 * scoreList[5]))));
            Point point6 = new Point(100 + 150, 150 + 50 + 15 * scoreList[6]);
            Point point7 = new Point(100 + 150 - ((int)(Math.Cos(Math.PI / 4) * (15 * scoreList[7]))), 50 + 150 + ((int)(Math.Sin(Math.PI / 4) * (15 * scoreList[7]))));

            #region 画区域
            Pen p1 = new Pen(Color.FromArgb(35, 150, 245), 1); //指定颜色和粗细

            g.DrawLine(p1, point0, point1);
            g.DrawLine(p1, point1, point2);
            g.DrawLine(p1, point2, point3);
            g.DrawLine(p1, point3, point4);
            g.DrawLine(p1, point4, point5);
            g.DrawLine(p1, point5, point6);
            g.DrawLine(p1, point6, point7);
            g.DrawLine(p1, point7, point0);
            #endregion

            #region 画圆点点
            Pen p2 = new Pen(Color.Green, 2f); //指定颜色和粗细
            Brush bs = new SolidBrush(Color.FromArgb(33, 150, 243));//此类不允许直接构造对象的，如果要构造对象只能用SolidBrush类为它专门构造对象，参数可以设置颜色
            g.DrawEllipse(p2, point0.X - 2, point0.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point0.X - 2, point0.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point1.X - 2, point1.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point1.X - 2, point1.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point2.X - 2, point2.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point2.X - 2, point2.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point3.X - 2, point3.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point3.X - 2, point3.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point4.X - 2, point4.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point4.X - 2, point4.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point5.X - 2, point5.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point5.X - 2, point5.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point6.X - 2, point6.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point6.X - 2, point6.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point7.X - 2, point7.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point7.X - 2, point7.Y - 2, 5, 5);//填充颜色

            #endregion

            #region 填写文字
            Brush bs2 = new SolidBrush(Color.Black);

            Font font = new Font("宋体", 10, FontStyle.Regular);

            // 绘制围绕点旋转的文本  
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Near;
            format.LineAlignment = StringAlignment.Center;

            Point fontPoint = new Point(1, 200);
            g.DrawString("身体动觉智能 " + scoreList[0].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(40, 80);
            g.DrawString("自然观察者智能 " + scoreList[1].ToString(), font, bs2, fontPoint, format);

            format.Alignment = StringAlignment.Center;
            fontPoint = new Point(250, 40);
            g.DrawString("人际交往智能 " + scoreList[2].ToString(), font, bs2, fontPoint, format);

            format.Alignment = StringAlignment.Near;

            fontPoint = new Point(350, 80);
            g.DrawString("自知内省智能 " + scoreList[3].ToString(), font, bs2, fontPoint, format);

            format.Alignment = StringAlignment.Near;
            fontPoint = new Point(400, 200);
            g.DrawString("语言文字智能 " + scoreList[4].ToString(), font, bs2, fontPoint, format);


            fontPoint = new Point(350, 320);
            g.DrawString("音乐节奏智能 " + scoreList[5].ToString(), font, bs2, fontPoint, format);

            format.Alignment = StringAlignment.Center;
            fontPoint = new Point(250, 370);
            g.DrawString("逻辑数学智能 " + scoreList[6].ToString(), font, bs2, fontPoint, format);

            format.Alignment = StringAlignment.Near;
            fontPoint = new Point(40, 320);
            g.DrawString("自然观察者智能 " + scoreList[7].ToString(), font, bs2, fontPoint, format);

            MemoryStream stream = new MemoryStream();
            bm.Save(stream, ImageFormat.Png);
            System.Drawing.Image image = System.Drawing.Image.FromStream(stream);
            image.Save("D://正八边形雷达图.png");

            #endregion
        }

        /// <summary>
        /// 中学生兴趣测评正六边形雷达图
        /// </summary>
        /// <param name="radius">正多边形外切圆的半径</param>
        /// <param name="sideCount">正多边形的边数</param>
        /// <param name="scoreList">正多边形的嵌套个数</param>
        public void Draw6RegularPoly(double radius, int[] scoreList)
        {
            double[] center = new double[] { 250.0, 200.0 };//正多边形外切圆的圆心固定

            int sideCount = 6;
            if (scoreList.Length != sideCount)
            {
                return;
            }
            // 每条边对应的圆心角角度，精确为浮点数。使用弧度制，360度角为2派
            double arc = 2 * Math.PI / sideCount;
            // 为多边形创建所有的顶点列表
            var pointList = new List<double[]>();
            System.Drawing.Point[] points = new System.Drawing.Point[6];
            Bitmap bm = new Bitmap(500, 400);
            Graphics g = Graphics.FromImage(bm);//画板
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            g.SmoothingMode = SmoothingMode.AntiAlias;  //使绘图质量最高，即消除锯齿
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.CompositingQuality = CompositingQuality.HighQuality;
            System.Drawing.Font f = new System.Drawing.Font("微软雅黑", 1);
            Pen p = new Pen(Color.Gray, 0.1f); //指定颜色和粗细
            //g.DrawLine(p, new System.Drawing.Point(0, 0), new System.Drawing.Point(800, 0));//X轴
            //g.DrawLine(p, new System.Drawing.Point(0, 0), new System.Drawing.Point(0, 800));//Y轴
            for (int j = 0; j < 3; j++)
            {
                radius -= 50;
                for (int i = 0; i < sideCount; i++)
                {
                    var curArc = arc * i; // 当前点对应的圆心角角度
                    double[] pt = new double[3];
                    // 就是简单的三角函数正余弦根据圆心角和半径算点坐标。这里都取整就行
                    pt[0] = center[0] + Math.Round((radius * Math.Cos(curArc)), 2);
                    pt[1] = center[1] + Math.Round((radius * Math.Sin(curArc)), 2);
                    pt[2] = 0;
                    pointList.Add(pt);

                    points[i] = new System.Drawing.Point((int)pt[0], (int)pt[1]);
                }
                SolidBrush barBrush1 = new SolidBrush(Color.Gray);
                SolidBrush barBrush2 = new SolidBrush(Color.FromArgb(0, 206, 209));


                g.DrawPolygon(p, points);//画多边形
                if (j == 0)//画对角线
                {
                    g.DrawLine(p, points[0], points[3]);
                    g.DrawLine(p, points[1], points[4]);
                    g.DrawLine(p, points[2], points[5]);
                }
            }
            float x = (float)150 / 90;
            int s = (int)(x * (90 - scoreList[0]));

            Point point0 = new Point(100 + s, 200);
            Point point1 = new Point(100 + 150 - (int)((Math.Cos(Math.PI / 3) * x * scoreList[1])), 50 + 150 - (int)(Math.Sin(Math.PI / 3) * x * scoreList[1]));
            Point point2 = new Point(250 + (int)((Math.Cos(Math.PI / 3) * x * scoreList[2])), 150 + 50 - (int)(Math.Sin(Math.PI / 3) * x * (scoreList[2])));
            Point point3 = new Point(250 + (int)(x * scoreList[3]), 200);
            Point point4 = new Point(250 + (int)((Math.Cos(Math.PI / 3) * x * scoreList[4])), 200 + (int)(Math.Sin(Math.PI / 3) * x * (scoreList[4])));
            Point point5 = new Point(100 + 150 - (int)((Math.Cos(Math.PI / 3) * x * scoreList[5])), 200 + (int)(Math.Sin(Math.PI / 3) * x * scoreList[5]));



            #region 画区域线
            Pen p1 = new Pen(Color.FromArgb(35, 150, 245), 1); //指定颜色和粗细
            g.DrawLine(p1, point0, point1);
            g.DrawLine(p1, point1, point2);
            g.DrawLine(p1, point2, point3);
            g.DrawLine(p1, point3, point4);
            g.DrawLine(p1, point4, point5);
            g.DrawLine(p1, point5, point0);
            #endregion

            #region 画圆点点
            Pen p2 = new Pen(Color.Green, 2f); //指定颜色和粗细
            Brush bs = new SolidBrush(Color.FromArgb(33, 150, 243));//此类不允许直接构造对象的，如果要构造对象只能用SolidBrush类为它专门构造对象，参数可以设置颜色
            g.DrawEllipse(p2, point0.X, point0.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point0.X, point0.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point1.X - 2, point1.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point1.X - 2, point1.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point2.X - 2, point2.Y, 5, 5);//画圆
            g.FillEllipse(bs, point2.X - 2, point2.Y, 5, 5);//填充颜色

            g.DrawEllipse(p2, point3.X - 2, point3.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point3.X - 2, point3.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point4.X - 2, point4.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point4.X - 2, point4.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p2, point5.X - 2, point5.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point5.X - 2, point5.Y - 2, 5, 5);//填充颜色


            #endregion

            #region 填写文字
            Brush bs2 = new SolidBrush(Color.Black);

            Font font = new Font("宋体", 10, FontStyle.Regular);

            // 绘制围绕点旋转的文本  
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Near;
            format.LineAlignment = StringAlignment.Center;

            Point fontPoint = new Point(1, 200);
            g.DrawString("实际型(R) " + scoreList[0].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(90, 60);
            g.DrawString("研究型(I) " + scoreList[1].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(320, 60);
            g.DrawString("艺术型(A) " + scoreList[2].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(400, 200);
            g.DrawString("社会型(S) " + scoreList[3].ToString(), font, bs2, fontPoint, format);


            fontPoint = new Point(320, 340);
            g.DrawString("企业型(E) " + scoreList[4].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(90, 340);
            g.DrawString("事务型(C) " + scoreList[5].ToString(), font, bs2, fontPoint, format);



            #endregion
            //Graphics graphics = g;
            //// 红色笔
            //Pen pen = new Pen(Color.Red, 5);
            //Rectangle rect = new Rectangle(0, 0, 200, 50);
            //// 用红色笔画矩形
            //graphics.DrawRectangle(pen, rect);
            //graphics.TranslateTransform(200, 0);
            //graphics.RotateTransform(90);
            //// 蓝色笔
            //pen.Color = Color.Blue;
            //// 用蓝色笔重新画平移之后的矩形
            //graphics.DrawRectangle(pen, rect);
            //graphics.Dispose();
            //pen.Dispose();

            MemoryStream stream = new MemoryStream();
            bm.Save(stream, ImageFormat.Png);
            System.Drawing.Image image = System.Drawing.Image.FromStream(stream);
            image.Save("D://正六边形雷达图.png");
        }

        /// <summary>
        /// 多元智能测评柱状图
        /// </summary>
        /// <param name="scoreList"></param>
        public void Columnar(int[] scoreList)
        {
            int sideCount = 8;
            if (scoreList.Length != sideCount)
            {
                return;
            }
            System.Drawing.Point[] points = new System.Drawing.Point[6];
            Bitmap bm = new Bitmap(680, 600);
            Graphics g = Graphics.FromImage(bm);//画板
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            g.SmoothingMode = SmoothingMode.AntiAlias;  //使绘图质量最高，即消除锯齿
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.CompositingQuality = CompositingQuality.HighQuality;
            System.Drawing.Font f = new System.Drawing.Font("微软雅黑", 1);



            Pen p1 = new Pen(Color.Gray, 0.05f); //指定颜色和粗细
            Pen p2 = new Pen(Color.FromArgb(220, 220, 220), 0.1f); //指定颜色和粗细
            #region 横线
            g.DrawLine(p1, new Point(0, 300), new Point(650, 300));
            #endregion
            #region 填写文字
            Brush bs3 = new SolidBrush(Color.Black);

            Font font = new Font("宋体", 10, FontStyle.Regular);

            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Near;
            format.LineAlignment = StringAlignment.Center;


            Point fontPoint = new Point(40, 320);
            g.DrawString("人际\r\n交往智能 ", font, bs3, fontPoint, format);

            fontPoint = new Point(120, 320);
            g.DrawString("自知\r\n内省智能 ", font, bs3, fontPoint, format);

            fontPoint = new Point(200, 320);
            g.DrawString("语言\r\n文字智能 ", font, bs3, fontPoint, format);

            fontPoint = new Point(280, 320);
            g.DrawString("音乐\r\n节奏智能 ", font, bs3, fontPoint, format);
            fontPoint = new Point(360, 320);
            g.DrawString("视觉\r\n空间智能 ", font, bs3, fontPoint, format);

            fontPoint = new Point(440, 320);
            g.DrawString("逻辑\r\n数学智能  ", font, bs3, fontPoint, format);
            fontPoint = new Point(520, 320);
            g.DrawString("身体\r\n动觉智能 ", font, bs3, fontPoint, format);
            fontPoint = new Point(600, 320);
            g.DrawString("自然\r\n观察者智能 ", font, bs3, fontPoint, format);


            #endregion
            #region 画区域线
            Brush bs0 = new SolidBrush(Color.FromArgb(99, 136, 244));//此类不允许直接构造对象的，如果要构造对象只能用SolidBrush类为它专门构造对象，参数可以设置颜色
            Brush bs1 = new SolidBrush(Color.FromArgb(150, 180, 250));//此类不允许直接构造对象的，如果要构造对象只能用SolidBrush类为它专门构造对象，参数可以设置颜色

            g.DrawRectangle(p2, 40, (10 - scoreList[0]) * 30, 40, (scoreList[0]) * 30);//画圆
            g.FillRectangle(bs0, 40, (10 - scoreList[0]) * 30, 40, (scoreList[0]) * 30);//填充颜色
            bs3 = new SolidBrush(Color.White);
            font = new Font("宋体", 10, FontStyle.Bold);
            fontPoint = new Point(47, (10 - scoreList[0]) * 30 + 20);
            g.DrawString(scoreList[0].ToString() + ".0", font, bs3, fontPoint, format);

            for (int i = 1; i < scoreList.Length; i++)
            {
                if (i % 2 == 0)
                {
                    g.DrawRectangle(p2, 80 * (i) + 40, (10 - scoreList[i]) * 30, 40, (scoreList[i]) * 30);//画圆
                    g.FillRectangle(bs0, 80 * (i) + 40, (10 - scoreList[i]) * 30, 40, (scoreList[i]) * 30);//填充颜色
                }
                else
                {
                    g.DrawRectangle(p2, 80 * (i) + 40, (10 - scoreList[i]) * 30, 40, (scoreList[i]) * 30);//画圆
                    g.FillRectangle(bs1, 80 * (i) + 40, (10 - scoreList[i]) * 30, 40, (scoreList[i]) * 30);//填充颜色
                }
                fontPoint = new Point(80 * (i) + 47, (10 - scoreList[i]) * 30 + 20);
                g.DrawString(scoreList[i].ToString() + ".0", font, bs3, fontPoint, format);

            }
            #endregion

            MemoryStream stream = new MemoryStream();
            bm.Save(stream, ImageFormat.Png);
            System.Drawing.Image image = System.Drawing.Image.FromStream(stream);
            image.Save("D://柱状图.png");
        }

        /// <summary>
        /// 中学生兴趣测评折线图
        /// </summary>
        /// <param name="scoreList"></param>
        public void BrokenLine(int[] scoreList)
        {
            int sideCount = 6;
            if (scoreList.Length != sideCount)
            {
                return;
            }
            System.Drawing.Point[] points = new System.Drawing.Point[6];
            Bitmap bm = new Bitmap(625, 605);
            Graphics g = Graphics.FromImage(bm);//画板
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            g.SmoothingMode = SmoothingMode.AntiAlias;  //使绘图质量最高，即消除锯齿
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.CompositingQuality = CompositingQuality.HighQuality;
            System.Drawing.Font f = new System.Drawing.Font("微软雅黑", 1);

            Point point0 = new Point(50, (100 - scoreList[0]) * 4);
            Point point1 = new Point(150, (100 - scoreList[1]) * 4);
            Point point2 = new Point(250, (100 - scoreList[2]) * 4);
            Point point3 = new Point(350, (100 - scoreList[3]) * 4);
            Point point4 = new Point(450, (100 - scoreList[4]) * 4);
            Point point5 = new Point(550, (100 - scoreList[5]) * 4);

            Pen p1 = new Pen(Color.Gray, 0.05f); //指定颜色和粗细
            Pen p2 = new Pen(Color.FromArgb(220, 220, 220), 0.1f); //指定颜色和粗细
            #region 横线

            g.DrawLine(p2, new Point(20, 40), new Point(620, 40));
            g.DrawLine(p2, new Point(20, 80), new Point(620, 80));
            g.DrawLine(p2, new Point(20, 120), new Point(620, 120));
            g.DrawLine(p2, new Point(20, 160), new Point(620, 160));
            g.DrawLine(p2, new Point(20, 200), new Point(620, 200));
            g.DrawLine(p2, new Point(20, 240), new Point(620, 240));
            g.DrawLine(p2, new Point(20, 280), new Point(620, 280));
            g.DrawLine(p2, new Point(20, 320), new Point(620, 320));
            g.DrawLine(p2, new Point(20, 360), new Point(620, 360));
            g.DrawLine(p1, new Point(20, 400), new Point(620, 400));

            g.DrawLine(p1, new Point(20, 380), new Point(20, 400));
            g.DrawLine(p1, new Point(120, 380), new Point(120, 400));
            g.DrawLine(p1, new Point(220, 380), new Point(220, 400));
            g.DrawLine(p1, new Point(320, 380), new Point(320, 400));
            g.DrawLine(p1, new Point(420, 380), new Point(420, 400));
            g.DrawLine(p1, new Point(520, 380), new Point(520, 400));
            g.DrawLine(p1, new Point(620, 380), new Point(620, 400));
            #endregion
            #region 填写文字
            Brush bs2 = new SolidBrush(Color.Black);

            Font font = new Font("宋体", 10, FontStyle.Regular);

            // 绘制围绕点旋转的文本  
            StringFormat format = new StringFormat();
            format.Alignment = StringAlignment.Near;
            format.LineAlignment = StringAlignment.Center;

            Point point01 = new Point(40, (100 - scoreList[0]) * 4 - 20);
            Point point11 = new Point(140, (100 - scoreList[1]) * 4 - 20);
            Point point21 = new Point(240, (100 - scoreList[2]) * 4 - 20);
            Point point31 = new Point(340, (100 - scoreList[3]) * 4 - 20);
            Point point41 = new Point(440, (100 - scoreList[4]) * 4 - 20);
            Point point51 = new Point(540, (100 - scoreList[5]) * 4 - 20);
            g.DrawString(scoreList[0].ToString(), font, bs2, point01);
            g.DrawString(scoreList[1].ToString(), font, bs2, point11);
            g.DrawString(scoreList[2].ToString(), font, bs2, point21);
            g.DrawString(scoreList[3].ToString(), font, bs2, point31);
            g.DrawString(scoreList[4].ToString(), font, bs2, point41);
            g.DrawString(scoreList[5].ToString(), font, bs2, point51);

            g.DrawString("90 ", font, bs2, new Point(0, 40), format);
            g.DrawString("80 ", font, bs2, new Point(0, 80), format);
            g.DrawString("70 ", font, bs2, new Point(0, 120), format);
            g.DrawString("60 ", font, bs2, new Point(0, 160), format);
            g.DrawString("50 ", font, bs2, new Point(0, 200), format);
            g.DrawString("40 ", font, bs2, new Point(0, 240), format);
            g.DrawString("30 ", font, bs2, new Point(0, 280), format);
            g.DrawString("20 ", font, bs2, new Point(0, 320), format);
            g.DrawString("10 ", font, bs2, new Point(0, 360), format);
            g.DrawString("0 ", font, bs2, new Point(0, 400), format);


            Point fontPoint = new Point(30, 420);
            g.DrawString("实际型(R) " + scoreList[0].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(130, 420);
            g.DrawString("研究型(I) " + scoreList[1].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(230, 420);
            g.DrawString("艺术型(A) " + scoreList[2].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(330, 420);
            g.DrawString("社会型(S) " + scoreList[3].ToString(), font, bs2, fontPoint, format);


            fontPoint = new Point(430, 420);
            g.DrawString("企业型(E) " + scoreList[4].ToString(), font, bs2, fontPoint, format);

            fontPoint = new Point(530, 420);
            g.DrawString("事务型(C) " + scoreList[5].ToString(), font, bs2, fontPoint, format);

            #endregion
            #region 画区域线
            Pen p3 = new Pen(Color.FromArgb(35, 150, 245), 1); //指定颜色和粗细
            g.DrawLine(p3, point0, point1);
            g.DrawLine(p3, point1, point2);
            g.DrawLine(p3, point2, point3);
            g.DrawLine(p3, point3, point4);
            g.DrawLine(p3, point4, point5);
            #endregion

            #region 画圆点点
            Pen p4 = new Pen(Color.Green, 2f); //指定颜色和粗细
            Brush bs = new SolidBrush(Color.FromArgb(33, 150, 243));//此类不允许直接构造对象的，如果要构造对象只能用SolidBrush类为它专门构造对象，参数可以设置颜色
            g.DrawEllipse(p4, point0.X, point0.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point0.X, point0.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p4, point1.X - 2, point1.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point1.X - 2, point1.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p4, point2.X - 2, point2.Y, 5, 5);//画圆
            g.FillEllipse(bs, point2.X - 2, point2.Y, 5, 5);//填充颜色

            g.DrawEllipse(p4, point3.X - 2, point3.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point3.X - 2, point3.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p4, point4.X - 2, point4.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point4.X - 2, point4.Y - 2, 5, 5);//填充颜色

            g.DrawEllipse(p4, point5.X - 2, point5.Y - 2, 5, 5);//画圆
            g.FillEllipse(bs, point5.X - 2, point5.Y - 2, 5, 5);//填充颜色


            #endregion

            MemoryStream stream = new MemoryStream();
            bm.Save(stream, ImageFormat.Png);
            System.Drawing.Image image = System.Drawing.Image.FromStream(stream);
            image.Save("D://折线图.png");
        }

        /// <summary>
        /// 使用外切圆的方法绘制一个正多边形
        /// </summary>
        /// <param name="stage">要绘制图形的设备</param>
        /// <param name="center">正多边形外切圆的圆心</param>
        /// <param name="radius">正多边形外切圆的半径</param>
        /// <param name="sideCount">正多边形的边数</param>
        public void DrawRegularPolyImg(Graphics stage, Point center, int radius, int sideCount)
        {
            // 多边形至少要有3条边，边数不达标就返回。
            if (sideCount < 3) return;
            // 每条边对应的圆心角角度，精确为浮点数。使用弧度制，360度角为2派
            double arc = 2 * Math.PI / sideCount;
            // 为多边形创建所有的顶点列表
            var pointList = new List<Point>();
            for (int i = 0; i < sideCount; i++)
            {
                var curArc = arc * i; // 当前点对应的圆心角角度
                var pt = new Point();
                // 就是简单的三角函数正余弦根据圆心角和半径算点坐标。这里都取整就行
                pt.X = center.X + (int)(radius * Math.Cos(curArc));
                pt.Y = center.Y + (int)(radius * Math.Sin(curArc));
                pointList.Add(pt);
            }
            // 在给出的场景中用黑笔把这个多边形画出来
            stage.DrawPolygon(Pens.Black, pointList.ToArray());
        }

    }
}
