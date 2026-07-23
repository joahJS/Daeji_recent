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
    public partial class PP002F00 : DevExpress.XtraEditors.XtraForm
    {
        public PP002F00()
        {
            InitializeComponent();
        }

        private string PROCEDURE_ID = "DP_PP002F00";
        public string _SCH_WORD;
        public string[] str = new string[100];
        public string Gu = "";

        public delegate void SendDataHandler(DataRow row);
        public event SendDataHandler DataRowSendEvent;

        public delegate void SendDataHandler2(string[] str);
        public event SendDataHandler2 DataRowSendEvent2;

        private void PP005F00_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            TxtFindWord.EditValue = _SCH_WORD;
            BtnRetr.PerformClick();
        }

        private void PP005F00_Shown(object sender, EventArgs e)
        {
            if (GridViewRetr.RowCount > 0)
            {
                GridViewRetr.Focus();
            }
            else if (GridViewRetr.RowCount == 0)
            {
                TxtFindWord.SelectAll();
                TxtFindWord.Focus();
            }
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> dicParams = new Dictionary<string, string>();
            dicParams.Add("CMD", "LIST1");
            dicParams.Add("FIND_IDX", CboFindIdx.SelectedIndex.ToString());
            dicParams.Add("FIND_WORD", TxtFindWord.EditValue?.ToString().Trim());

            GridRetr.DataSource = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
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
                XtraMessageBox.Show("사용자리스트의 내용이 존재하지 않습니다.");
                return;
            }

            if (!Gu.Equals("1"))
            {
                DataRowSendEvent(GridViewRetr.GetFocusedDataRow());
                DialogResult = DialogResult.OK;
                Dispose();
            }
            else
            {
                DataRowSendEvent2(str);
                DialogResult = DialogResult.OK;
                Dispose();
            }

        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Dispose();
        }

        private void PP005F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                BtnClose.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnApply.PerformClick();
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnApply.PerformClick();
        }
        int i = 0;
        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
                BtnApply.PerformClick();

            if (Gu.Equals("1"))
            {
                str[i] += GridViewRetr.GetFocusedRowCellValue(GridColUsrNm)?.ToString();
                i++;
            }
        }
    }
}