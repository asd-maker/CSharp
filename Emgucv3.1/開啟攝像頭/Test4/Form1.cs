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

namespace Test4
{
    public partial class Form1 : Form
    {
        private Capture _capture = null;
        private Mat _frame;
        public Form1()
        {
            InitializeComponent();

        }
        private void ProcessFrame(object sender, EventArgs e)
        {
            if (_capture != null && _capture.Ptr != IntPtr.Zero)
            {
                _capture.Retrieve(_frame, 0);
                imageBox1.Image = _frame;//imageBox1显示控件
            }
        }
        private void btnOpenCapture_Click(object sender, EventArgs e)
        {
            _capture = new Capture();
            _capture.ImageGrabbed += ProcessFrame;
            _frame = new Mat();
            if (_capture != null) _capture.Start();//摄像头开启
        }
        private void btnStop_Click(object sender, EventArgs e)
        {
            _capture.Stop();//摄像头关闭
        }

    }
}
