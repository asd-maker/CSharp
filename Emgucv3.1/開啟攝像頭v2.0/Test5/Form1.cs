using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Test5
{
    public partial class Form1 : Form
    {
        private Capture _capture = null;
        private Mat _frame;
        private Capture _capture1 = null;
        private Mat _frame1;
        public Form1()
        {
            InitializeComponent();

        }
        private void ProcessFrame(object sender, EventArgs e)
        {
            if (_capture != null && _capture.Ptr != IntPtr.Zero)
            {
                _capture.Retrieve(_frame, 0);
                pictureBox1.Image = _frame.Bitmap;//imageBox1显示控件
                //pictureBox1.Image= _capture.QueryFrame().Bitmap;
 
            }
        }
        private void ProcessFrame1(object sender, EventArgs e)
        {
            if (_capture1 != null && _capture1.Ptr != IntPtr.Zero)
            {
                _capture1.Retrieve(_frame1, 0);
                imageBox2.Image = _frame1;//imageBox1显示控件
            }
        }
        private void btnOpenCapture_Click(object sender, EventArgs e)
        {
            _capture = new Capture(0);
            _capture1 = new Capture(1);
            _capture1.SetCaptureProperty(CapProp.FrameWidth, 450);
            _capture1.SetCaptureProperty(CapProp.FrameHeight, 380);
            _capture.ImageGrabbed += ProcessFrame;
            _capture1.ImageGrabbed += ProcessFrame1;
            _frame = new Mat();
            _frame1 = new Mat();
            if (_capture != null) _capture.Start();//摄像头开启
            if (_capture1 != null) _capture1.Start();//摄像头开启
        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            pictureBox1.Image.Save(@"C:\Users\lei.l.zhao\Desktop\1.jpg");
            _capture.Stop();//摄像头关闭
        }
    }
}
