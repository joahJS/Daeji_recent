using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;
using System.Collections.Generic;
using DevExpress.XtraPrinting.Drawing;

namespace AccAdm
{
    public partial class RptMesuring2 : DevExpress.XtraReports.UI.XtraReport
    {
        public RptMesuring2()
        {
            InitializeComponent();
            
        }

        public RptMesuring2(DataRow row, Dictionary<string, Image> dicImage)
        {
            InitializeComponent();

            LblDealerNm.Text = row["DEALER_NM"]?.ToString();
            if (row["DEALER_NM"]?.ToString().Length > 14)
            {
                LblDealerNm.Font = new System.Drawing.Font("굴림체", 10F, System.Drawing.FontStyle.Bold);
            }
            LblJ_Date.Text = row["J_DATE"]?.ToString();
            LblKeraType.Text = row["KERATYPE"]?.ToString();
            LblSun.Text = row["SUN"]?.ToString();
            LblGrade1.Text = row["GUBUN1"]?.ToString();
            LblJ_BNum.Text = row["J_BNUM"]?.ToString();
            LblGumsu.Text = row["EMP_NM"]?.ToString();
            LblTotWeight.Text = row["TOT_WEIGHT"]?.ToString();
            LblEmptyWeight.Text = row["EMPTY_WEIGHT"]?.ToString();
            LblAcptWeight.Text = row["ACTL_WEIGHT"]?.ToString();
            LblGrade2.Text = row["GUBUN2"]?.ToString();
            LblLoss.Text = row["LOSS"]?.ToString();
            LblGumsuBigo.Text = row["J_STATE"]?.ToString();

            string sKERATYPE = row["KERATYPE"]?.ToString();
            if (sKERATYPE.Equals("입고"))
            {
                LblFirstTime.Text = row["FIRSTTIME"]?.ToString();
                LblSecondTime.Text = row["SECONDTIME"]?.ToString();
            }
            else
            {
                LblFirstTime.Text = row["SECONDTIME"]?.ToString();
                LblSecondTime.Text = row["FIRSTTIME"]?.ToString();
            }

            ////XrPicBox1.Sizing = DevExpress.XtraPrinting.ImageSizeMode.StretchImage;
            Pic1.ImageSource = new ImageSource(dicImage["1"]);
            Pic2.ImageSource = new ImageSource(dicImage["2"]);
            Pic3.ImageSource = new ImageSource(dicImage["3"]);
            Pic4.ImageSource = new ImageSource(dicImage["4"]);
            Pic5.ImageSource = new ImageSource(dicImage["5"]);
            Pic6.ImageSource = new ImageSource(dicImage["6"]);
        }
    }
}
