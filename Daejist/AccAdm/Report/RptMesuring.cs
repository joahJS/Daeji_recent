using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using DevExpress.XtraPrinting.Drawing;
using System.Collections.Generic;
using System.Data;

namespace AccAdm
{
    public partial class RptMesuring : DevExpress.XtraReports.UI.XtraReport
    {
        public RptMesuring()
        {
            InitializeComponent();

            XrLblDealerNm.DataBindings.Add("Text", null, "DEALER_NM");
            XrLblJDate.DataBindings.Add("Text", null, "J_DATE");
            XrLblKeraType.DataBindings.Add("Text", null, "KERATYPE");
            XrLblSun.DataBindings.Add("Text", null, "SUN");
            XrLblGubun1.DataBindings.Add("Text", null, "GUBUN1");
            XrLblJBnum.DataBindings.Add("Text", null, "J_BNUM");
            XrLblGumsuSerial.DataBindings.Add("Text", null, "EMP_NM");
            XrLblTotWeight.DataBindings.Add("Text", null, "TOT_WEIGHT");
            XrLblEmptyWeight.DataBindings.Add("Text", null, "EMPTY_WEIGHT");
            XrLblActualWeight.DataBindings.Add("Text", null, "ACTL_WEIGHT");
            XrLblGubun1_2.DataBindings.Add("Text", null, "GUBUN2");
            XrLblLoss.DataBindings.Add("Text", null, "LOSS");
            XrLblGumsubigo.DataBindings.Add("Text", null, "J_STATE");
        }

        public RptMesuring(DataRow row, Dictionary<string, Image> dicImage)
        {
            InitializeComponent();
            
            XrLblDealerNm.Text = row["DEALER_NM"]?.ToString();
            if (row["DEALER_NM"]?.ToString().Length > 14)
            {
                XrLblDealerNm.Font = new System.Drawing.Font("굴림체", 10F, System.Drawing.FontStyle.Bold);
            }
            XrLblJDate.Text = row["J_DATE"]?.ToString();
            XrLblKeraType.Text = row["KERATYPE"]?.ToString();
            XrLblSun.Text = row["SUN"]?.ToString();
            XrLblGubun1.Text = row["GUBUN1"]?.ToString();
            XrLblJBnum.Text = row["J_BNUM"]?.ToString();
            XrLblGumsuSerial.Text = row["EMP_NM"]?.ToString();
            XrLblTotWeight.Text = row["TOT_WEIGHT"]?.ToString();
            XrLblEmptyWeight.Text = row["EMPTY_WEIGHT"]?.ToString();
            XrLblActualWeight.Text = row["ACTL_WEIGHT"]?.ToString();
            XrLblGubun1_2.Text = row["GUBUN2"]?.ToString();
            XrLblLoss.Text = row["LOSS"]?.ToString();
            XrLblGumsubigo.Text = row["J_STATE"]?.ToString();
            //XrPicBox1.Sizing = DevExpress.XtraPrinting.ImageSizeMode.StretchImage;
            XrPicBox1.ImageSource = new ImageSource(dicImage["1"]);
            XrPicBox2.ImageSource = new ImageSource(dicImage["2"]);
            XrPicBox3.ImageSource = new ImageSource(dicImage["3"]);
            XrPicBox4.ImageSource = new ImageSource(dicImage["4"]);
            XrPicBox5.ImageSource = new ImageSource(dicImage["5"]);
            XrPicBox6.ImageSource = new ImageSource(dicImage["6"]);
        }
    }
}
