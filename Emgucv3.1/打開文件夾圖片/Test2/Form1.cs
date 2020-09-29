using Emgu.CV;
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

namespace Test2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            //加载图像
            var dialog = new OpenFileDialog();
            dialog.Filter = "图片(*.jpg/*.png/*.gif/*.bmp)|*.jpg;*.png;*.gif;*.bmp";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var filename = dialog.FileName;
                Image<Bgr, Byte> image;
                image = new Image<Bgr, byte>(@filename);
                imageBox1.Image = image;
            }
        }
    }
}
