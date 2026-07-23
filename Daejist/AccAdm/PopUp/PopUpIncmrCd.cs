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
    public partial class PopUpIncmrCd : DevExpress.XtraEditors.XtraForm
    {
        public PopUpIncmrCd()
        {
            InitializeComponent();
        }

        public string sIncmrNmByVal { get; set; }
        public DataRow drResult { get; set; }

        private void PopUpIncmrCd_Load(object sender, EventArgs e)
        {
            TxtIncmrNm.Text = sIncmrNmByVal;

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sIncmrNm = TxtIncmrNm.Text.Trim();

            if(sIncmrNm.Length == 0)
            {
                MessageBox.Show("조회할 소득자명을 입력하세요!");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT IFNULL(CONCAT(LEFT(INCMR_IDT_NO, 4), '**-*******'),  CONCAT(LEFT(PP_NO, 4), '******')) AS INCMR_IDT_NO ");
            strSql.AppendLine("      , INCMR_NM ");
            strSql.AppendLine("      , CO_NM ");
            strSql.AppendLine("      , TEL_NO ");
            strSql.AppendLine("      , EMAIL ");
            //strSql.AppendLine("      , BZRG_NO ");
            //strSql.AppendLine("      , BZPLC_LOC ");
            //strSql.AppendLine("      , RESID_NTN ");
            //strSql.AppendLine("      , ZIP ");
            //strSql.AppendLine("      , ZIP_SEQ ");
            //strSql.AppendLine("      , ADDR ");
            //strSql.AppendLine("      , DTL_ADDR ");
            //strSql.AppendLine("      , BIZ_NM ");
            //strSql.AppendLine("      , BANK_CD ");
            //strSql.AppendLine("      , BANK_ACNT_NO ");
            //strSql.AppendLine("      , ACNT_HOLDER ");
            //strSql.AppendLine("      , USE_YN ");
            //strSql.AppendLine("      , NOTE ");
            strSql.AppendLine("      , INCMR_CD ");
            strSql.AppendLine("   FROM ACC_INCM_CD A ");
            strSql.AppendLine("  WHERE 1 = 1 ");
            strSql.AppendLine("    AND INCMR_NM LIKE CONCAT('%" + sIncmrNm + "', '%') ");
            strSql.AppendLine("  ORDER BY INCMR_CD ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("조회된 결과가 없습니다.\r\n소득자명을 다시 확인하세요!");
                return;
            }
            else if (dt.Rows.Count == 1)
            {
                drResult = dt.Rows[0];
                this.DialogResult = DialogResult.OK;
                return;
            }

            GridIncmrCd.DataSource = dt;
            GridViewIncmrCd.FocusedRowHandle = 0;
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

        private void CloseOk()
        {
            drResult = GridViewIncmrCd.GetFocusedDataRow();

            this.DialogResult = DialogResult.OK;
        }

        private void TxtIncmrNm_Leave(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void GridViewIncmrCd_DoubleClick(object sender, EventArgs e)
        {
            CloseOk();
        }

        private void PopUpIncmrCd_KeyDown(object sender, KeyEventArgs e)
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

        private void GridIncmrCd_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewIncmrCd_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}