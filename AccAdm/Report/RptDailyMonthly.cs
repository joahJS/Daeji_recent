using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using DevExpress.XtraReports.UI;

namespace AccAdm
{
    public partial class RptDailyMonthly : DevExpress.XtraReports.UI.XtraReport
    {
        public RptDailyMonthly()
        {
            InitializeComponent();
            TableCellShredderOut.DataBindings.Add("Text", null, "SWEIGHT");
            TableCellShredderTaget.DataBindings.Add("Text", null, "STAGET");
            TableCellShredderMinus.DataBindings.Add("Text", null, "SMINER");

            TableCellWeightTaget.DataBindings.Add("Text", null, "WTAGET");
            TableCellWeightOut.DataBindings.Add("Text", null, "WWEIGHT");
            TableCellWeightMinus.DataBindings.Add("Text", null, "WMINER");
        }
    }
}
