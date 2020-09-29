using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Emgu.CV;
using Emgu.CV.Structure;
using Emgucv;

namespace Emgucv
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Image<Bgr, byte> image = new Image<Bgr, byte>(320, 640, new Bgr(0, 0, 255));
            Image<Bgr, byte> image1 = new Image<Bgr, byte>(320, 640);
            imageBox1.Image = image;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog op = new OpenFileDialog();
            if (op.ShowDialog() == DialogResult.OK)
            {
                Mat img = new Mat(op.FileName, Emgu.CV.CvEnum.LoadImageType.AnyColor);
                imageBox1.Image = img;
            }
            //點，线，矩形，園，橢圓
            MCvPoint3D32f mCvPoint3D32F = new MCvPoint3D32f(0, 0, 0);
            PointF x = new PointF(0, 0);
            PointF y = new PointF(1, 1);
            LineSegment2DF lineSegment2DF = new LineSegment2DF(x, y);

            //顏色
            Rgb rgb = new Rgb(Color.Red);
            Rgb red = new Rgb(255, 0, 0);
            //類型轉換
            Bitmap bitmap = new Bitmap(640, 480);
            Image<Bgr, byte> image = new Image<Bgr, byte>(640, 480);
            Mat mat = new Mat();
            op = new OpenFileDialog();
            if (op.ShowDialog() == DialogResult.OK)
            {
                Bitmap bitmap_new = new Bitmap(op.FileName);
                Image<Bgr, byte> image_new = new Image<Bgr, byte>(op.FileName);
                Mat mat_new = new Mat(op.FileName, Emgu.CV.CvEnum.LoadImageType.AnyColor);

                pictureBox1.Image = bitmap_new;
                pictureBox1.Image = image_new.ToBitmap();
                //pictureBox1.Image = mat_new.Bitmap;
                pictureBox1.Image = mat_new.ToImage<Bgr, byte>().Bitmap;

                imageBox1.Image = new Image<Bgr, byte>(bitmap_new);
                imageBox1.Image = image_new;
                imageBox1.Image = mat_new;

            }
            Image<Bgr, byte> img = new Image<Bgr, byte>(200,200,new Bgr(255,255,255));
            Rectangle rectangle = new Rectangle(new Point(80,80),new Size(40,40));
            CircleF circleF = new CircleF(new PointF(100,100),40);
            string str = "I LOVE EmguCV";
            Point str_location = new Point(0,30);
            img.Draw(rectangle,new Bgr(0,255,0),2);
            img.Draw(circleF,new Bgr(0,0,255),3);
            img.Draw(str,str_location,Emgu.CV.CvEnum.FontFace.HersheyComplexSmall,1,new Bgr(0,255,0),3);
            imageBox1.Image = img;





        }
    }
}
