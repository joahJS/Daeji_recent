using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace AccAdm
{
    public partial class PopEtcCostDtl : DevExpress.XtraEditors.XtraForm
    {
        public PopEtcCostDtl()
        {
            InitializeComponent();
        }

        public DataRow rowInfo;

        private void PopEtcCostDtl_Load(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("ETC_DEALER_NM1");
            dt.Columns.Add("ETC_REMARK1");
            dt.Columns.Add("ETC_COST1");
            dt.Columns.Add("ETC_DEALER_NM2");
            dt.Columns.Add("ETC_REMARK2");
            dt.Columns.Add("ETC_COST2");

            DataRow row = dt.NewRow();

            row["ETC_DEALER_NM1"] = rowInfo["ETC_DEALER_NM1"];
            row["ETC_REMARK1"] = rowInfo["ETC_REMARK1"];
            row["ETC_COST1"] = double.Parse(rowInfo["ETC_COST1"].ToString()).ToString("#,#");
            row["ETC_DEALER_NM2"] = rowInfo["ETC_DEALER_NM2"];
            row["ETC_REMARK2"] = rowInfo["ETC_REMARK2"];
            row["ETC_COST2"] = double.Parse(rowInfo["ETC_COST2"].ToString()).ToString("#,#");

            dt.Rows.Add(row);
            GridRetr.DataSource = dt;
        }

        private void PopEtcCostDtl_Shown(object sender, EventArgs e)
        {
            BtnClose.Focus();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void PopEtcCostDtl_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
                BtnClose.PerformClick();
        }
        
    }
}