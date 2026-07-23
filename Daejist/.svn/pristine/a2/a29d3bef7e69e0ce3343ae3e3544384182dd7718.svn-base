using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;

namespace AccAdm
{
    public partial class AllReportIM01 : DevExpress.XtraEditors.XtraForm
    {
        public byte[] ImgData;

        public AllReportIM01()
        {
            InitializeComponent();
        }

        private void AllReportIM01_Load(object sender, EventArgs e)
        {
            IMAGERE();
        }
        private void IMAGERE()
        {
            MemoryStream ms = new MemoryStream(ImgData);
            Image returnImage = Image.FromStream(ms);
            //PicInsa.Image = ResizeImage(returnImage, 300, 200);
            PicInsa.Image = returnImage;
            
        }

        private void PicInsa_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}