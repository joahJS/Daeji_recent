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
    public partial class PopUpDeptCd : DevExpress.XtraEditors.XtraForm
    {
        public PopUpDeptCd()
        {
            InitializeComponent();
        }

        public string sDeptCdByVal { get;  set; }
        public string sDeptNmByVal { get; set; }
        public DataRow drResult { get; set; }

        private void PopUpDeptCd_Load(object sender, EventArgs e)
        {
            TxtDeptCd.Text = sDeptCdByVal;
            TxtDeptNm.Text = sDeptNmByVal;

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sDeptCd = TxtDeptCd.Text.Trim();
            string sDeptNm = TxtDeptNm.Text.Trim();

            if(sDeptCd.Length == 0 && sDeptNm.Length == 0)
            {
                MessageBox.Show("조회할 계정코드 또는 계정명을 입력하세요!");
                TxtDeptNm.Focus();
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT DEPT_CD ");
            strSql.AppendLine("      , DEPT_NM  ");
            strSql.AppendLine("      , UP_DEPT_CD ");
            strSql.AppendLine("      , (SELECT X.DEPT_NM FROM daejierp.ACC_DEPT_CD X WHERE X.DEPT_CD = A.UP_DEPT_CD) AS UP_DEPT_NM ");
            strSql.AppendLine("      , DEPT_GB  ");
            strSql.AppendLine("   FROM daejierp.ACC_DEPT_CD A ");
            strSql.AppendLine("  WHERE USE_YN = 'Y' ");

            if (sDeptCd.Length == 0)
            {
                strSql.AppendLine("    AND DEPT_NM LIKE CONCAT('%', '" + sDeptNm + "', '%') ");
            }
            else
            {
                strSql.AppendLine("    AND DEPT_CD LIKE CONCAT('" + sDeptCd + "', '%') ");
            }

            strSql.AppendLine("  ORDER BY DEPT_CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("조회된 결과가 없습니다.\r\n부서코드 또는 부서명을 다시 확인하세요!");
                TxtDeptNm.Focus();
                return;
            }
            else if (dt.Rows.Count == 1)
            {
                drResult = dt.Rows[0];
                this.DialogResult = DialogResult.OK;
                return;
            }

            GridDeptCd.DataSource = dt;
            GridViewDeptCd.FocusedRowHandle = 0;
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
            drResult = GridViewDeptCd.GetFocusedDataRow();

            this.DialogResult = DialogResult.OK;
        }

        private void TxtAccCd_Leave(object sender, EventArgs e)
        {
            if (!TxtDeptCd.Text.Equals("")) TxtDeptNm.Text = "";
            BtnRetr_Click(null, null);
        }

        private void TxtAccNm_Leave(object sender, EventArgs e)
        {
            if (!TxtDeptNm.Text.Equals("")) TxtDeptCd.Text = "";
            BtnRetr_Click(null, null);
        }

        private void PopUpDeptCd_KeyDown(object sender, KeyEventArgs e)
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

        private void GridDeptCd_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewDeptCd_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}