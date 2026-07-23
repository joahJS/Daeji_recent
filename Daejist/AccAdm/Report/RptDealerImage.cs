using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Collections.Generic;
using System.Data;
using DevExpress.XtraPrinting.Drawing;
using System.IO;
using System.Drawing.Drawing2D;

namespace AccAdm
{
    public partial class RptDealerImage : DevExpress.XtraReports.UI.XtraReport
    {
        public RptDealerImage()
        {
            InitializeComponent();
        }

        public static Image byteArrayToImage(byte[] byteArrayIn)
        {
            Image returnImage = null;
            try
            {
                System.IO.MemoryStream ms = new MemoryStream(byteArrayIn, 0, byteArrayIn.Length);
                ms.Write(byteArrayIn, 0, byteArrayIn.Length);
                returnImage = Image.FromStream(ms, true);//Exception occurs here
            }
            catch { }
            return returnImage;
        }


        public RptDealerImage(DataTable dt1)
        {
            InitializeComponent();
            byte[] Img = Convert.IsDBNull(dt1.Rows[0]["IMIMAGE"]) ? null : (byte[])dt1.Rows[0]["IMIMAGE"];
            
            Image JLine_Image = byteArrayToImage(Img);
            JLine_Image = ResizeImage(JLine_Image, 750, 1100);
            if (Img != null)
            {
                xrPictureBox1.ImageSource = new ImageSource(JLine_Image);
            }
        }

        private Image ResizeImage(Image image, int width, int height)
        {
            // 새로운 이미지 생성
            Bitmap resizedImage = new Bitmap(width, height);

            // 그래픽 개체 생성
            using (Graphics graphics = Graphics.FromImage(resizedImage))
            {
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;

                // 이미지를 새로운 크기로 그립니다.
                graphics.DrawImage(image, 0, 0, width, height);
            }

            return resizedImage;
        }
    }
}
