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
    public partial class PopUpGradeCd : DevExpress.XtraEditors.XtraForm
    {
        public PopUpGradeCd()
        {
            InitializeComponent();
        }

        public string sGradeCd { get; set; }
        public string sGradeNm { get; set; }
        public DataRow drResult { get; set; }

        private void PopUpGradeCd_Load(object sender, EventArgs e)
        {
            TxtGrade.EditValue = sGradeNm;

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sGrade = TxtGrade.Text == null ? string.Empty : TxtGrade.Text;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.J_SERIAL ");
            strSql.AppendLine("      , A.GUBUN1 ");
            strSql.AppendLine("      , A.DAEGUBUN ");
            strSql.AppendLine("   FROM JAJAE A");
            strSql.AppendLine("  WHERE GUBUN1 LIKE '%" + sGrade + "%' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("조회된 결과가 없습니다.\r\n 등급명을 다시 확인하세요!");
                return;
            }
            else if (dt.Rows.Count == 1)
            {
                drResult = dt.Rows[0];
                this.DialogResult = DialogResult.OK;
                return;
            }
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

        private void TxtGrade_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void TxtGrade_Leave(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            drResult = null;
            this.DialogResult = DialogResult.Cancel;
            Close();
        }

        private void PopUpGradeCd_KeyDown(object sender, KeyEventArgs e)
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