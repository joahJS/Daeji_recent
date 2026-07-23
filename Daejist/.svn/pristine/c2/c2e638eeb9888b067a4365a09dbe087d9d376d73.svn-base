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
    public partial class PopUpMgmtCd : DevExpress.XtraEditors.XtraForm
    {
        public PopUpMgmtCd()
        {
            InitializeComponent();
        }

        public string sMgmtGb { get; set; }
        public string sMgmtCdByVal { get;  set; }
        public string sMgmtNmByVal { get; set; }
        public DataRow drResult { get; set; }

        private void PopUpMgmtCd_Load(object sender, EventArgs e)
        {
            TxtMgmtCd.Text = sMgmtCdByVal;
            TxtMgmtNm.Text = sMgmtNmByVal;

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sMgmtCd = TxtMgmtCd.Text.Trim();
            string sMgmtNm = TxtMgmtNm.Text.Trim();

            //if(sMgmtCd.Length == 0 && sMgmtNm.Length == 0)
            //{
            //    MessageBox.Show("조회할 관리항목코드 또는 관리항목명을 입력하세요!");
            //    //TxtMgmtNm.Focus();
            //    return;
            //}

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT CD ");
            strSql.AppendLine("      , NM ");
            strSql.AppendLine("   FROM V_MGMT_CD A ");
            strSql.AppendLine("  WHERE MGMT_GB =  '" + sMgmtGb +"' ");

            if (sMgmtCd.Length == 0)
            {
                strSql.AppendLine("    AND NM LIKE CONCAT('%', '" + sMgmtNm + "', '%') ");
            }
            else
            {
                strSql.AppendLine("    AND CD LIKE CONCAT('%', '" + sMgmtCd + "', '%') ");
            }

            strSql.AppendLine("  ORDER BY CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("조회된 결과가 없습니다.\r\n관리항목코드 또는 관리항목명을 다시 확인하세요!");
                //TxtMgmtNm.Focus();
                return;
            }
            else if (dt.Rows.Count == 1)
            {
                drResult = dt.Rows[0];
                this.DialogResult = DialogResult.OK;
            }

            GridMgmtCd.DataSource = dt;
            GridViewMgmtCd.FocusedRowHandle = 0;
        }

        private void BtnConFirm_Click(object sender, EventArgs e)
        {
            CloseOk();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void CloseOk()
        {
            drResult = GridViewMgmtCd.GetFocusedDataRow();
            this.DialogResult = DialogResult.OK;
        }


        private void TxtMgmtCd_Leave(object sender, EventArgs e)
        {
            if (!TxtMgmtCd.Text.Equals("")) TxtMgmtNm.Text = "";
            BtnRetr_Click(null, null);
        }

        private void TxtMgmtNm_Leave(object sender, EventArgs e)
        {
            if (!TxtMgmtNm.Text.Equals("")) TxtMgmtCd.Text = "";
            BtnRetr_Click(null, null);
        }

        private void GridViewMgmtCd_DoubleClick(object sender, EventArgs e)
        {
            CloseOk();
        }

        private void PopUpMgmtCd_KeyDown(object sender, KeyEventArgs e)
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

        private void GridMgmtCd_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewMgmtCd_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}