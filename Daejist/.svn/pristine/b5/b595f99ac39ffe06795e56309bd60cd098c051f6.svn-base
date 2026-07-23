using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;

namespace AccAdm
{
    public partial class RptAcc : DevExpress.XtraReports.UI.XtraReport
    {
        public RptAcc()
        {
            InitializeComponent();

        }

        public RptAcc(DataTable dtM, DataTable dtD)
        {
            InitializeComponent();
            BindingData(dtD);

            PV_IDT_NO.Text = dtM.Rows[0]["PV_IDT_NO"]?.ToString();
            PV_CVNAM.Text = dtM.Rows[0]["PV_CVNAM"]?.ToString();
            PV_REP_NM.Text = dtM.Rows[0]["PV_REP_NM"]?.ToString();
            PV_ADDR.Text = dtM.Rows[0]["PV_ADDR"]?.ToString();
            PV_BIZ_NM.Text = dtM.Rows[0]["PV_BIZ_NM"]?.ToString();
            PV_TYPE_NM.Text = dtM.Rows[0]["PV_TYPE_NM"]?.ToString();

            BY_IDT_NO.Text = dtM.Rows[0]["BY_IDT_NO"]?.ToString();
            BY_CVNAM.Text = dtM.Rows[0]["BY_CVNAM"]?.ToString();
            BY_REP_NM.Text = dtM.Rows[0]["BY_REP_NM"]?.ToString();
            BY_ADDR.Text = dtM.Rows[0]["BY_ADDR"]?.ToString();
            BY_BIZ_NM.Text = dtM.Rows[0]["BY_BIZ_NM"]?.ToString();
            BY_TYPE_NM.Text = dtM.Rows[0]["BY_TYPE_NM"]?.ToString();

            
        }
        private void BindingData(DataTable dtD)
        {
            
            xrTDATE.DataBindings.Add("Text", dtD, "TDATE");
            xrATEXT.DataBindings.Add("Text", dtD, "ATEXT");
            xrACAMT.DataBindings.Add("Text", dtD, "ACAMT");
            xrADAMT.DataBindings.Add("Text", dtD, "ADAMT");
            xrJJAMT.DataBindings.Add("Text", dtD, "JJAMT");
            
        }
    }
}
