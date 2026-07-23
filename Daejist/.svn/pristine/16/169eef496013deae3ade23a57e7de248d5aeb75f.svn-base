using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;
using System.Data;

namespace AccAdm
{
    public partial class RptTaxBill2 : DevExpress.XtraReports.UI.XtraReport
    {
        public RptTaxBill2()
        {
            InitializeComponent();
        }

        public RptTaxBill2(DataTable dt)
        {
            InitializeComponent();

            PV_IDT_NO.Text = dt.Rows[0]["PV_IDT_NO"]?.ToString();
            PV_CVNAM.Text = dt.Rows[0]["PV_CVNAM"]?.ToString();
            PV_REP_NM.Text = dt.Rows[0]["PV_REP_NM"]?.ToString();
            PV_ADDR.Text = dt.Rows[0]["PV_ADDR"]?.ToString();
            PV_BIZ_NM.Text = dt.Rows[0]["PV_BIZ_NM"]?.ToString();
            PV_TYPE_NM.Text = dt.Rows[0]["PV_TYPE_NM"]?.ToString();

            BY_IDT_NO.Text = dt.Rows[0]["BY_IDT_NO"]?.ToString();
            BY_CVNAM.Text = dt.Rows[0]["BY_CVNAM"]?.ToString();
            BY_REP_NM.Text = dt.Rows[0]["BY_REP_NM"]?.ToString();
            BY_ADDR.Text = dt.Rows[0]["BY_ADDR"]?.ToString();
            BY_BIZ_NM.Text = dt.Rows[0]["BY_BIZ_NM"]?.ToString();
            BY_TYPE_NM.Text = dt.Rows[0]["BY_TYPE_NM"]?.ToString();

            TDATE.Text = dt.Rows[0]["TDATE"]?.ToString();
            TOT_AMT.Text = dt.Rows[0]["TOT_AMT"]?.ToString();
            TOT_VAT.Text = dt.Rows[0]["TOT_VAT"]?.ToString();
            RK.Text = dt.Rows[0]["RK"]?.ToString();

            MM1.Text = dt.Rows[0]["MM1"]?.ToString();
            DD1.Text = dt.Rows[0]["DD1"]?.ToString();
            ITNM1.Text = dt.Rows[0]["ITNM1"]?.ToString();
            SPEC1.Text = dt.Rows[0]["SPEC1"]?.ToString();
            TWGT1.Text = dt.Rows[0]["TWGT1"]?.ToString();
            COST1.Text = dt.Rows[0]["COST1"]?.ToString();
            AMT1.Text = dt.Rows[0]["TAMT1"]?.ToString();
            VAT1.Text = dt.Rows[0]["TVAT1"]?.ToString();

            MM2.Text = dt.Rows[0]["MM2"]?.ToString();
            DD2.Text = dt.Rows[0]["DD2"]?.ToString();
            ITNM2.Text = dt.Rows[0]["ITNM2"]?.ToString();
            SPEC2.Text = dt.Rows[0]["SPEC2"]?.ToString();
            TWGT2.Text = dt.Rows[0]["TWGT2"]?.ToString();
            COST2.Text = dt.Rows[0]["COST2"]?.ToString();
            AMT2.Text = dt.Rows[0]["TAMT2"]?.ToString();
            VAT2.Text = dt.Rows[0]["TVAT2"]?.ToString();

            MM3.Text = dt.Rows[0]["MM3"]?.ToString();
            DD3.Text = dt.Rows[0]["DD3"]?.ToString();
            ITNM3.Text = dt.Rows[0]["ITNM3"]?.ToString();
            SPEC3.Text = dt.Rows[0]["SPEC3"]?.ToString();
            TWGT3.Text = dt.Rows[0]["TWGT3"]?.ToString();
            COST3.Text = dt.Rows[0]["COST3"]?.ToString();
            AMT3.Text = dt.Rows[0]["TAMT3"]?.ToString();
            VAT3.Text = dt.Rows[0]["TVAT3"]?.ToString();

            MM4.Text = dt.Rows[0]["MM4"]?.ToString();
            DD4.Text = dt.Rows[0]["DD4"]?.ToString();
            ITNM4.Text = dt.Rows[0]["ITNM4"]?.ToString();
            SPEC4.Text = dt.Rows[0]["SPEC4"]?.ToString();
            TWGT4.Text = dt.Rows[0]["TWGT4"]?.ToString();
            COST4.Text = dt.Rows[0]["COST4"]?.ToString();
            AMT4.Text = dt.Rows[0]["TAMT4"]?.ToString();
            VAT4.Text = dt.Rows[0]["TVAT4"]?.ToString();

            SUM_AMT.Text = dt.Rows[0]["SUM_AMT"]?.ToString();

            string sPayGb = dt.Rows[0]["PAYGB"]?.ToString();
            if (sPayGb.Equals("1"))
            {
                CASH.Text = "●";
            }
            else if (sPayGb.Equals("2"))
            {
                CHECK.Text = "●";
            }
            else if (sPayGb.Equals("3"))
            {
                NOTE.Text = "●";
            }
            else if (sPayGb.Equals("4") || sPayGb.Equals("5"))
            {
                MISU.Text = "●";
            }

            //RST_MSG.Text = "위 금액을 청구함";

            //string sTaxGu = dt.Rows[0]["TAXGU"]?.ToString();
            //if (sTaxGu.Equals("S"))
            //{
            //    PV_IDT_NO.Text = dt.Rows[0]["PV_IDT_NO"]?.ToString();
            //    PV_CVNAM.Text = dt.Rows[0]["PV_CVNAM"]?.ToString();
            //    PV_REP_NM.Text = dt.Rows[0]["PV_REP_NM"]?.ToString();
            //    PV_ADDR.Text = dt.Rows[0]["PV_ADDR"]?.ToString();
            //    PV_BIZ_NM.Text = dt.Rows[0]["PV_BIZ_NM"]?.ToString();
            //    PV_TYPE_NM.Text = dt.Rows[0]["PV_TYPE_NM"]?.ToString();

            //    BY_IDT_NO.Text = dt.Rows[0]["BY_IDT_NO"]?.ToString();
            //    BY_CVNAM.Text = dt.Rows[0]["BY_CVNAM"]?.ToString();
            //    BY_REP_NM.Text = dt.Rows[0]["BY_REP_NM"]?.ToString();
            //    BY_ADDR.Text = dt.Rows[0]["BY_ADDR"]?.ToString();
            //    BY_BIZ_NM.Text = dt.Rows[0]["BY_BIZ_NM"]?.ToString();
            //    BY_TYPE_NM.Text = dt.Rows[0]["BY_TYPE_NM"]?.ToString();

            //    TDATE.Text = dt.Rows[0]["TDATE"]?.ToString();
            //    TOT_AMT.Text = dt.Rows[0]["TOT_AMT"]?.ToString();
            //    TOT_VAT.Text = dt.Rows[0]["TOT_VAT"]?.ToString();
            //    RK.Text = dt.Rows[0]["RK"]?.ToString();

            //    MM1.Text = dt.Rows[0]["MM1"]?.ToString();
            //    DD1.Text = dt.Rows[0]["DD1"]?.ToString();
            //    ITNM1.Text = dt.Rows[0]["ITNM1"]?.ToString();
            //    SPEC1.Text = dt.Rows[0]["SPEC1"]?.ToString();
            //    TWGT1.Text = dt.Rows[0]["TWGT1"]?.ToString();
            //    COST1.Text = dt.Rows[0]["COST1"]?.ToString();
            //    AMT1.Text = dt.Rows[0]["AMT1"]?.ToString();
            //    VAT1.Text = dt.Rows[0]["VAT1"]?.ToString();

            //    MM2.Text = dt.Rows[0]["MM2"]?.ToString();
            //    DD2.Text = dt.Rows[0]["DD2"]?.ToString();
            //    ITNM2.Text = dt.Rows[0]["ITNM2"]?.ToString();
            //    SPEC2.Text = dt.Rows[0]["SPEC2"]?.ToString();
            //    TWGT2.Text = dt.Rows[0]["TWGT2"]?.ToString();
            //    COST2.Text = dt.Rows[0]["COST2"]?.ToString();
            //    AMT2.Text = dt.Rows[0]["AMT2"]?.ToString();
            //    VAT2.Text = dt.Rows[0]["VAT2"]?.ToString();

            //    MM3.Text = dt.Rows[0]["MM3"]?.ToString();
            //    DD3.Text = dt.Rows[0]["DD3"]?.ToString();
            //    ITNM3.Text = dt.Rows[0]["ITNM3"]?.ToString();
            //    SPEC3.Text = dt.Rows[0]["SPEC3"]?.ToString();
            //    TWGT3.Text = dt.Rows[0]["TWGT3"]?.ToString();
            //    COST3.Text = dt.Rows[0]["COST3"]?.ToString();
            //    AMT3.Text = dt.Rows[0]["AMT3"]?.ToString();
            //    VAT3.Text = dt.Rows[0]["VAT3"]?.ToString();

            //    MM4.Text = dt.Rows[0]["MM4"]?.ToString();
            //    DD4.Text = dt.Rows[0]["DD4"]?.ToString();
            //    ITNM4.Text = dt.Rows[0]["ITNM4"]?.ToString();
            //    SPEC4.Text = dt.Rows[0]["SPEC4"]?.ToString();
            //    TWGT4.Text = dt.Rows[0]["TWGT4"]?.ToString();
            //    COST4.Text = dt.Rows[0]["COST4"]?.ToString();
            //    AMT4.Text = dt.Rows[0]["AMT4"]?.ToString();
            //    VAT4.Text = dt.Rows[0]["VAT4"]?.ToString();

            //    SUM_AMT.Text = dt.Rows[0]["SUM_AMT"]?.ToString();

            //    string sPayGb = dt.Rows[0]["PAYGB"]?.ToString();
            //    RST_MSG.Text = "위 금액을 청구함";


            //}
            //else if (sTaxGu.Equals("P"))
            //{

            //}

        }
    }
}
