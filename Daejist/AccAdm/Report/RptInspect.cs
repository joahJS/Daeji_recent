using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;
using System.Collections.Generic;

namespace AccAdm
{
    public partial class RptInspect : DevExpress.XtraReports.UI.XtraReport
    {
        public RptInspect()
        {
            InitializeComponent();
        }

        /*
         * dic : 상단 출력 값
         * dt1 : 스크랩
         * dt2 : 슈레더
        */
        public RptInspect(Dictionary<string, string> dic, DataTable dt1, DataTable dt2)
        {
            InitializeComponent();

            DATE.Text = dic["DATE"];
            RETURN1.Text = dic["RETURN1"];
            MINUS1.Text = dic["MINUS1"];
            MINUS_WGT1.Text = dic["MINUS_WGT1"];
            
            RETURN2.Text = dic["RETURN2"];
            MINUS2.Text = dic["MINUS2"];
            MINUS_WGT2.Text = dic["MINUS_WGT2"];

            DEALER1.DataBindings.Add("Text", dt1, "DEALER1");
            IN_WGT1.DataBindings.Add("Text", dt1, "IN_WGT1");
            MINUS_WGT_SUM1.DataBindings.Add("Text", dt1, "MINUS_WGT_SUM1");
            IN_CNT1.DataBindings.Add("Text", dt1, "IN_CNT1");
            MINUS_CNT1.DataBindings.Add("Text", dt1, "MINUS_CNT1");
            ADMT1.DataBindings.Add("Text", dt1, "ADMT1");
            ITNL1.DataBindings.Add("Text", dt1, "ITNL1");
            OPN1.DataBindings.Add("Text", dt1, "OPN1");

            DEALER2.DataBindings.Add("Text", dt2, "DEALER2");
            IN_WGT2.DataBindings.Add("Text", dt2, "IN_WGT2");
            MINUS_WGT_SUM2.DataBindings.Add("Text", dt2, "MINUS_WGT_SUM2");
            IN_CNT2.DataBindings.Add("Text", dt2, "IN_CNT2");
            MINUS_CNT2.DataBindings.Add("Text", dt2, "MINUS_CNT2");
            ADMT2.DataBindings.Add("Text", dt2, "ADMT2");
            ITNL2.DataBindings.Add("Text", dt2, "ITNL2");
            OPN2.DataBindings.Add("Text", dt2, "OPN2");


        }
    }
}
