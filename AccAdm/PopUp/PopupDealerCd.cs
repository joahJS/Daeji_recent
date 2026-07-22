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
    public partial class PopUpDealerCd : DevExpress.XtraEditors.XtraForm
    {
        public PopUpDealerCd()
        {
            InitializeComponent();
        }

        public string sIdtNoByVal { get; set; }
        public string sDealerNmByVal { get; set; }
        public DataRow drResult { get; set; }

        private void PopupDealerCd_Load(object sender, EventArgs e)
        {
            TxtIdtNo.EditValue = sIdtNoByVal;
            TxtDealerNm.EditValue = sDealerNmByVal;

            BtnRetr_Click(null, null);
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            string sIdtNo = TxtIdtNo.EditValue?.ToString().Replace("-", "").Replace("_", "");
            string sDealerNm = TxtDealerNm.Text.Trim();

            //if (sIdtNo.Length == 0 && sDealerNm.Length == 0)
            //{
            //    XtraMessageBox.Show("조회할 사업자등록번호 또는 거래처명을 입력하세요!");
            //    return;
            //}

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.DEALER_CD ");
            strSql.AppendLine("      , A.DEALER_NM ");
            strSql.AppendLine("      , A.DEALER_GB ");
            strSql.AppendLine("      , A.IDT_NO ");
            strSql.AppendLine("   FROM ACC_DEALER_CD A ");
            strSql.AppendLine("  WHERE 1 = 1");

            if (sIdtNo?.Length != 0)
            {
                strSql.AppendLine("    AND IDT_NO LIKE CONCAT('%" + sIdtNo + "', '%') ");
            }
            if (sDealerNm.Length != 0)
            {
                strSql.AppendLine("    AND DEALER_NM LIKE CONCAT('%', '" + sDealerNm + "', '%') ");
            }

            strSql.AppendLine("  ORDER BY A.DEALER_NM");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("조회된 결과가 없습니다.\r\n사업자등록번호 또는 거래처명을 다시 확인하세요!");
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

        private void TxtDealerNm_Leave(object sender, EventArgs e)
        {
            if (!TxtDealerNm.Text.Equals("")) TxtIdtNo.Text = "";
            BtnRetr_Click(null, null);
        }

        private void TxtIdtNo_Leave(object sender, EventArgs e)
        {
            if (!TxtIdtNo.Text.Equals("")) TxtDealerNm.Text = "";
            BtnRetr_Click(null, null);
        }

        private void TxtDealerNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void TxtIdtNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            drResult = null;
            this.DialogResult = DialogResult.Cancel;
        }

        private void PopUpDealerCd_KeyDown(object sender, KeyEventArgs e)
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