using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;

namespace AccAdm
{
    public partial class RptDealerCd : DevExpress.XtraReports.UI.XtraReport
    {
        public RptDealerCd()
        {
            InitializeComponent();
            XtbcdDealerNm.DataBindings.Add("Text", null, "DEALER_NM");
            Xtbcdchrg_nm.DataBindings.Add("Text", null, "CHRG_NM");
            XtbcdDEALER_GB.DataBindings.Add("Text", null, "DEALER_GB");
            XtbcdREP_NM.DataBindings.Add("Text", null, "REP_NM");
            XtbcdBIZ_NM.DataBindings.Add("Text", null, "BIZ_NM");
            XtbcdADDR.DataBindings.Add("Text", null, "ADDR");
            XtbcdEMAIL.DataBindings.Add("Text", null, "EMAIL");
            XtbcdFAX.DataBindings.Add("Text", null, "FAX");
        }
    }
}
