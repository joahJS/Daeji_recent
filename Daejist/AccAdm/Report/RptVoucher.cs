using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;

namespace AccAdm
{
    public partial class RptVoucher : DevExpress.XtraReports.UI.XtraReport
    {
        public RptVoucher()
        {
            InitializeComponent();
        }

        public RptVoucher(DataTable dt, ArrayList arr)
        {
            InitializeComponent();
            EMP_NM.Text = arr[0].ToString();
            PRINT_DT.Text = arr[1].ToString();
            DT_FROM_TO.Text = arr[2].ToString();

            /*
             * 2020-11-29 수정변경
             * 현업요청
             */
            XR_TDATE.DataBindings.Add("Text", null, "TDATE");
            XR_DEALER_NM.DataBindings.Add("Text", null, "DEALER_NM");
            XR_IDT_NO.DataBindings.Add("Text", null, "IDT_NO");
            XR_RMK.DataBindings.Add("Text", null, "RMK");
            XR_TRSUM2.DataBindings.Add("Text", null, "TRSUM");
            XR_AMT.DataBindings.Add("Text", null, "AMT"); 
            XR_DIFF.DataBindings.Add("Text", null, "DIFF");//XR_DIFF
            XR_BANK.DataBindings.Add("Text", null, "BANK");
            XR_ACNT_NO.DataBindings.Add("Text", null, "BANK_ACNT_NO");
            XR_ACNT_HOLDER.DataBindings.Add("Text", null, "ACNT_HOLDER");

            //XR_TDATE.DataBindings.Add("Text", null, "TDATE");
            //XR_DEALER_NM.DataBindings.Add("Text", null, "DEALER_NM");
            //XR_IDT_NO.DataBindings.Add("Text", null, "IDT_NO");
            //XR_RMK.DataBindings.Add("Text", null, "RMK");
            //XR_AMT.DataBindings.Add("Text", null, "AMT");
            //XR_TRSUM2.DataBindings.Add("Text", null, "TRSUM");
            //XR_DIFF.DataBindings.Add("Text", null, "DIFF");
            //XR_BANK.DataBindings.Add("Text", null, "BANK");
            //XR_ACNT_NO.DataBindings.Add("Text", null, "BANK_ACNT_NO");
            //XR_ACNT_HOLDER.DataBindings.Add("Text", null, "ACNT_HOLDER");

            XR_TOT_AMT.Text = arr[3].ToString(); 
        } 
    }
}
