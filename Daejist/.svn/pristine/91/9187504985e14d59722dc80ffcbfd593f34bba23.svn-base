using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;
using System.Collections.Generic;

namespace AccAdm
{
    public partial class RptInOutSlip : DevExpress.XtraReports.UI.XtraReport
    {
        public RptInOutSlip()
        {
            InitializeComponent();
        }

        public RptInOutSlip(List<string> List, DataTable dt)
        {
            InitializeComponent();

            DateTime dtPars = DateTime.Parse(List[5]);
            LblDate.Text = dtPars.ToString("yyyy 년 MM 월 dd 일");
            LblInOutGB.Text = List[0];

            CellSEQ.DataBindings.Add("Text", null, "SEQ");
            CellACNAM.DataBindings.Add("Text", null, "ACNAM");
            /*
             * 2022-01-25
             * 사업자번호 -> 거래처명으로 변경
             * (현업요청)
             */
            //CellIDT_NO.DataBindings.Add("Text", null, "IDT_NO");
            CellIDT_NO.DataBindings.Add("Text", null, "CVNAM");
            CellATEXT.DataBindings.Add("Text", null, "ATEXT");
            CellACAMT.DataBindings.Add("Text", null, "ACAMT");
            CellADAMT.DataBindings.Add("Text", null, "ADAMT");

            LblEmpNm.Text = List[1];
            LblPrintEmpNm.Text = List[2];
            LblSumACAMT.Text = List[3];
            LblSumADAMT.Text = List[4];
            LblNow.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

    }
}
