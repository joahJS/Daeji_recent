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

using ComLib;

namespace AccAdm
{
    public partial class PopUpAccCd : DevExpress.XtraEditors.XtraForm
    {
        public PopUpAccCd()
        {
            InitializeComponent();
        }

        public string sAccCdByVal { get;  set; }
        public string sAccNmByVal { get; set; }
        public DataRow drResult { get; set; }

        private void PopUpAccCd_Load(object sender, EventArgs e)
        {
            TxtAccCd.Text = sAccCdByVal;
            TxtAccNm.Text = sAccNmByVal;

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sAccCd = TxtAccCd.Text.Trim();
            string sAccNm = TxtAccNm.Text.Trim();

            if(sAccCd.Length == 0 && sAccNm.Length == 0)
            {
                MessageBox.Show("조회할 계정코드 또는 계정명을 입력하세요!");
                //TxtAccNm.Focus();
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT ACC_CD ");
            strSql.AppendLine("      , ACC_NM ");
            strSql.AppendLine("      , UP_ACC_CD ");
            strSql.AppendLine("      , (SELECT X.ACC_NM FROM ACC_ACC_CD X WHERE X.ACC_CD = A.UP_ACC_CD) AS UP_ACC_NM ");
            strSql.AppendLine("      , DBCR_GB  ");
            strSql.AppendLine("      , PYBC_YN ");
            strSql.AppendLine("      , EVDN_ESSN_YN ");
            strSql.AppendLine("      , DEALER_YN ");
            strSql.AppendLine("      , DEBT_MGMT_GB1 AS DEBT_MGMT_GB ");
            strSql.AppendLine("      , (SELECT X.COM_NM FROM COM_BASE_CD X WHERE X.CD_GB = 'MGMT_GB' AND X.COM_CD = A.DEBT_MGMT_GB1) AS DEBT_MGMT_GB_NM ");
            strSql.AppendLine("      , CRDT_MGMT_GB1 AS CRDT_MGMT_GB ");
            strSql.AppendLine("      , (SELECT X.COM_NM FROM COM_BASE_CD X WHERE X.CD_GB = 'MGMT_GB' AND X.COM_CD = A.CRDT_MGMT_GB1) AS CRDT_MGMT_GB_NM ");
            strSql.AppendLine("      , CPTL_CD ");
            strSql.AppendLine("      , RPLC_ACC_CD ");
            strSql.AppendLine("   FROM ACC_ACC_CD A ");
            strSql.AppendLine("  WHERE 1 = 1 ");

            if (sAccCd.Length == 0)
            {
                strSql.AppendLine("    AND ACC_NM LIKE CONCAT('%', '" + sAccNm + "', '%') ");
            }
            else
            {
                strSql.AppendLine("    AND ACC_CD LIKE CONCAT('" + sAccCd + "', '%') ");
            }

            strSql.AppendLine("  ORDER BY ACC_CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("조회된 결과가 없습니다.\r\n계정코드 또는 계정명을 다시 확인하세요!");
                //TxtAccNm.Focus();
                return;
            }
            else if (dt.Rows.Count == 1)
            {
                drResult = dt.Rows[0];
                this.DialogResult = DialogResult.OK;
                return;
            }

            GridAccCd.DataSource = dt;
            GridViewAccCd.FocusedRowHandle = 0;
        }

        private void BtnConFirm_Click(object sender, EventArgs e)
        {
            CloseOk();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            drResult = null;
            this.DialogResult = DialogResult.Cancel;
        }

        private void GridViewAccCd_DoubleClick(object sender, EventArgs e)
        {
            CloseOk();
        }

        private void CloseOk()
        {
            drResult = GridViewAccCd.GetFocusedDataRow();

            this.DialogResult = DialogResult.OK;
        }

        private void TxtAccCd_Leave(object sender, EventArgs e)
        {
            if (!TxtAccCd.Text.Equals("")) TxtAccNm.Text = "";
            BtnRetr_Click(null, null);
        }

        private void TxtAccNm_Leave(object sender, EventArgs e)
        {
            if (!TxtAccNm.Text.Equals("")) TxtAccCd.Text = "";
            BtnRetr_Click(null, null);
        }

        private void PopUpAccCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnConFirm_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {

            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
            }
        }

        private void GridAccCd_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewAccCd_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}