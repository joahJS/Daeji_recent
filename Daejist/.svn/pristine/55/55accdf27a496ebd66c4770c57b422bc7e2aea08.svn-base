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
    public partial class AC02001F06 : DevExpress.XtraEditors.XtraForm
    {
        public AC02001F06()
        {
            InitializeComponent();
        }

        public string FIND_WORD;
        public AC02001F02 PARENT_FORM;
        public AC14001F02 P_AC14001F02;

        private void AC02001F06_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(FIND_WORD))
                TxtFindWord.EditValue = FIND_WORD;

            FmMainToolBar2._FontSetting.SetGridView(GridViewRetr);
            BtnRetr_Click(null, null);
        }

        private void AC02001F06_Shown(object sender, EventArgs e)
        {
            if(GridViewRetr.RowCount > 0)
            {
                GridViewRetr.Focus();
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sFindSubject = CboFindSubject.SelectedIndex.ToString();
            string sFindWord = TxtFindWord.EditValue?.ToString();
            if (string.IsNullOrEmpty(sFindSubject))
            {
                XtraMessageBox.Show("찾을 항목을 선택하세요.");
                CboFindSubject.Focus();
                return;
            }

            GetDeptInfo(sFindSubject, sFindWord);
        }
        
        private void GetDeptInfo(string sFindSubject, string sFindWord)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine("   SELECT DEPT_CD ");
            strSql.AppendLine("        , DEPT_NM ");
            strSql.AppendLine("        , ( SELECT X1.DEPT_NM FROM ACC_DEPT_CD X1 WHERE X1.DEPT_CD = A.UP_DEPT_CD ) AS UP_DEPT_NM ");
            strSql.AppendLine("     FROM ACC_DEPT_CD A  ");
            strSql.AppendLine("    WHERE A.USE_YN = 'Y' ");
            strSql.AppendLine("      AND (('" + sFindSubject + "' = '0' AND DEPT_CD LIKE '%" + sFindWord + "%') ");
            strSql.AppendLine("           OR ");
            strSql.AppendLine("           ('" + sFindSubject + "' = '1' AND DEPT_NM LIKE '%" + sFindWord + "%')) ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;
            if(dt.Rows.Count > 0)
            {
                GridViewRetr.Focus();
            }
            else
            {
                TxtFindWord.SelectAll();
                TxtFindWord.Focus();
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if(PARENT_FORM != null)
            {
                PARENT_FORM.DR_DEPT_INFO = GridViewRetr.GetFocusedDataRow();
                DialogResult = DialogResult.OK;
            }
            else if(P_AC14001F02 != null)
            {
                P_AC14001F02.DR_DEPT_INFO = GridViewRetr.GetFocusedDataRow();
                DialogResult = DialogResult.OK;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void GridViewRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
                BtnApply_Click(null, null);
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
                BtnApply_Click(null, null);
        }
    }
}