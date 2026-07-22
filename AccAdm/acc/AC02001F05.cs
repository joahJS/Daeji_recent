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
using DevExpress.XtraGrid.Views.Grid;
using ComLib;

namespace AccAdm
{
    public partial class AC02001F05 : DevExpress.XtraEditors.XtraForm
    {
        public AC02001F05()
        {
            InitializeComponent();
        }

        public string BankCd;

        public AC02001F04 P_AC02001F04;


        private void AC02001F05_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(BankCd))
                TxtBankCd.EditValue = BankCd;

            FmMainToolBar2._FontSetting.SetGridView(GridViewRetr);
            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sBankCd = TxtBankCd.EditValue?.ToString();
            string sBankNm = TxtBankNm.EditValue?.ToString();

            GridRetr.DataSource = GetBankInfo(sBankCd, sBankNm);
        }

        private DataTable GetBankInfo(string sBankCd, string sBankNm)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.COM_CD AS BANK_CD ");
            strSql.AppendLine(" 	 , A.COM_NM AS BANK_NM ");
            strSql.AppendLine("   FROM COM_BASE_CD A   ");
            strSql.AppendLine("  WHERE CD_GB = 'BANK_CD' ");
            strSql.AppendLine("    AND A.COM_CD LIKE '%" + sBankCd + "%'");
            strSql.AppendLine("    AND A.COM_NM LIKE '%" + sBankNm + "%' ");
            
            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }

        private void BtnSelect_Click(object sender, EventArgs e)
        {
            AdmitSeletedRow();
            DialogResult = DialogResult.OK;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void AdmitSeletedRow()
        {
            DataRow row = GridViewRetr.GetFocusedDataRow();

            if (P_AC02001F04 != null)
                P_AC02001F04.DrBankInfo = row;
        }

        #region[KeyDown Event]

        private void AC02001F05_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSelect_Click(null, null);
            }
            else if(e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void TxtDealerCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void TxtDealerNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }

        #endregion[KeyDown Event]
        
        #region[GridView Design]

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                AdmitSeletedRow();
                DialogResult = DialogResult.OK;
            }
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }


        #endregion[GridView Design]
    }
}