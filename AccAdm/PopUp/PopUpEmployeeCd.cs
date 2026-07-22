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
using DevExpress.XtraEditors.Repository;

namespace AccAdm
{
    public partial class PopUpEmployeeCd : DevExpress.XtraEditors.XtraForm
    {
        public PopUpEmployeeCd()
        {
            InitializeComponent();
        }

        public string sEmpId { get; set; }
        public string sEmpNm { get; set; }
        public DataRow drResult { get; set; }

        private void PopUpEmployeeCd_Load(object sender, EventArgs e)
        {
            DataTable dtDeptNm = GetLookUpData("1", "Y", "Y");

            RepositoryItemGridLookUpEdit repoGLkupDeptNm = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(repoGLkupDeptNm, dtDeptNm, GridRetr, GridColDeptCd, "CD", "NM", "");

            TxtEmpNm.EditValue = sEmpNm;

            BtnRetr_Click(null, null);
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '공용' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , A.DEPT_CD AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  WHERE USE_YN = 'Y' ");
            }

            if (sParam.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            return dt;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sEmpNm = TxtEmpNm.EditValue?.ToString().Trim();

            if (string.IsNullOrEmpty(sEmpNm))
            {
                XtraMessageBox.Show("조회할 직원명을 입력하세요!");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.EMP_NM ");
            strSql.AppendLine("      , A.DEPT_CD ");
            strSql.AppendLine("      , A.ENTRANCE_YMD ");
            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
            strSql.AppendLine("  WHERE EMPL_GB = 'Y' ");
            if ( !sEmpNm.Equals(string.Empty))
            {
                strSql.AppendLine("  AND EMP_NM LIKE CONCAT('" + sEmpNm + "', '%') ");
            }
            strSql.AppendLine("  ORDER BY EMP_ID");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("조회된 결과가 없습니다.\r\n직원명을 다시 확인하세요!");
                return;
            }
            else if (dt.Rows.Count == 1)
            {
                drResult = dt.Rows[0];
                this.DialogResult = DialogResult.OK;
                return;
            }

            GridRetr.DataSource = dt;
            GridViewRetr.FocusedRowHandle = 0;

        }

        private void BtnConfirm_Click(object sender, EventArgs e)
        {
            CloseOk();
        }

        private void GridRetr_DoubleClick(object sender, EventArgs e)
        {
            CloseOk();
        }

        private void CloseOk()
        {
            drResult = GridViewRetr.GetFocusedDataRow();

            this.DialogResult = DialogResult.OK;
        }

        private void TxtEmpNm_Leave(object sender, EventArgs e)
        {
            if (!TxtEmpNm.Text.Equals("")) return;
            BtnRetr_Click(null, null);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            drResult = null;
            this.DialogResult = DialogResult.Cancel;
        }

        private void TxtEmpNm_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void PopUpEmployeeCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnConfirm_Click(null, null);
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

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}