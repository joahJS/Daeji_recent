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
    public partial class PopUpSchAcnt : DevExpress.XtraEditors.XtraForm
    {
        public PopUpSchAcnt()
        {
            InitializeComponent();
        }

        public delegate void SendDataHandler(DataRow row);
        public event SendDataHandler DataRowSendEvent;

        public string _FINDWORD = "";

        private void PopUpSchAcnt_Load(object sender, EventArgs e)
        {

            if (int.TryParse(_FINDWORD, out int result))
            {
                CboFindObj.SelectedIndex = 0;
            }
            else
            {
                CboFindObj.SelectedIndex = 1;
            }

            TxtFindWord.EditValue = _FINDWORD;

            BtnRetr.PerformClick();
        }

        private void PopUpSchAcnt_Shown(object sender, EventArgs e)
        {
            if (GridViewRetr.RowCount > 0 && !string.IsNullOrEmpty(_FINDWORD))
            {
                GridViewRetr.Focus();
            }
            else if (GridViewRetr.RowCount == 0)
            {
                TxtFindWord.SelectAll();
                TxtFindWord.Focus();
            }
            else
            {
                TxtFindWord.Focus();
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            int sFINDIDX = CboFindObj.SelectedIndex;
            string sFINDWORD = TxtFindWord.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine("SELECT A1.ACNT_CD                ");
            strSql.AppendLine("     , A1.BANK_ACNT_NO           ");
            strSql.AppendLine("     , A1.BANK_CD                ");
            strSql.AppendLine("     , B1.COM_NM AS BANK_NM      ");
            strSql.AppendLine("     , A1.PSSEQ                  ");
            strSql.AppendLine("     , A1.GEJAGB                 ");
            strSql.AppendLine("     , B2.COM_NM AS GGBNM        ");
            strSql.AppendLine("     , A1.ACC_CD                 ");
            strSql.AppendLine("     , C1.ACNAM AS ACNAM         ");
            strSql.AppendLine("     , A1.RPLC_ACC_CD            ");
            strSql.AppendLine("     , C2.ACNAM AS RPLCNM        ");
            strSql.AppendLine("     , A1.FINANCE_GOODS_NM       ");
            strSql.AppendLine("  FROM ACC_ACNT_CD A1            ");
            strSql.AppendLine("  LEFT JOIN COM_BASE_CD B1       ");
            strSql.AppendLine("    ON A1.BANK_CD = B1.COM_CD    ");
            strSql.AppendLine("   AND B1.CD_GB = 'BANK_CD'      ");
            strSql.AppendLine("  LEFT JOIN COM_BASE_CD B2       ");
            strSql.AppendLine("    ON A1.GEJAGB = B2.COM_CD     ");
            strSql.AppendLine("   AND B2.CD_GB = 'GEJAGB'       ");
            strSql.AppendLine("  LEFT JOIN ACMSTF C1            ");
            strSql.AppendLine("    ON A1.ACC_CD = C1.ACCOD      ");
            strSql.AppendLine("  LEFT JOIN ACMSTF C2            ");
            strSql.AppendLine("    ON A1.RPLC_ACC_CD = C2.ACCOD ");
            strSql.AppendLine(" WHERE 1 = 1                     ");
            if(!string.IsNullOrEmpty(sFINDWORD) && sFINDIDX == 0)//계좌번호
            {
                strSql.AppendLine("AND A1.BANK_ACNT_NO LIKE '%" + sFINDWORD + "%'");
            }
            if(!string.IsNullOrEmpty(sFINDWORD) && sFINDIDX == 1)//은행명
            {
                strSql.AppendLine("AND B1.COM_NM LIKE '%" + sFINDWORD + "%'");
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr.DataSource = dt;

            if (GridViewRetr.RowCount > 0)
                GridViewRetr.Focus();
            else if (GridViewRetr.RowCount == 0)
            {
                TxtFindWord.SelectAll();
                TxtFindWord.Focus();
            }
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (GridViewRetr.RowCount == 0)
            {
                XtraMessageBox.Show("계좌데이터가 존재하지 않습니다.");
                return;
            }

            DataRowSendEvent(GridViewRetr.GetFocusedDataRow());
            DialogResult = DialogResult.OK;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void PopUpSchAcnt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                BtnClose.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnAply.PerformClick();
            else if (e.KeyCode == Keys.Enter)
                SendKeys.Send("{TAB}");
        }

        private void GridViewRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnAply.PerformClick();
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }


        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
                BtnAply.PerformClick();
        }

        private void TxtFindWord_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                BtnRetr.Focus();
            }
        }

        private void CboFindObj_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TxtFindWord.Focus();
            }
        }
    }
}
