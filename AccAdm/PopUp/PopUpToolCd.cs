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
    public partial class PopUpToolCd : DevExpress.XtraEditors.XtraForm
    {
        public PopUpToolCd()
        {
            InitializeComponent();
        }

        public string sToolCd { get; set; }
        public string sToolNm { get; set; }
        public DataRow drResult { get; set; }

        private void PopUpToolCd_Load(object sender, EventArgs e)
        {
            TxtEquipToolNm.EditValue = sToolNm;

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sToolName = TxtEquipToolNm.EditValue == null ? string.Empty : TxtEquipToolNm.EditValue.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MG_NO ");
            strSql.AppendLine("      , A.TOOL_NM ");
            strSql.AppendLine("      , A.MG_DEPT ");
            strSql.AppendLine("      , A.MODEL_NM ");
            strSql.AppendLine("   FROM EQUIP_TOOL_MGT A ");
            strSql.AppendLine("  WHERE TOOL_NM LIKE '%" + sToolName + "%' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("조회된 결과가 없습니다.\r\n설비명을 다시 확인하세요!");
                return;
            }
            else if (dt.Rows.Count == 1)
            {
                drResult = dt.Rows[0];
                this.DialogResult = DialogResult.OK;
                return;
            }

            GridRetr.DataSource = dt;
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
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            drResult = null;
            this.DialogResult = DialogResult.Cancel;
        }

        private void TxtEquipToolNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void TxtEquipToolNm_Leave(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void PopUpToolCd_KeyDown(object sender, KeyEventArgs e)
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