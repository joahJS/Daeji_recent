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
    public partial class PP003F00 : DevExpress.XtraEditors.XtraForm
    {
        public PP003F00()
        {
            InitializeComponent();
        }
        private string PROCEDURE_ID = "DP_CM005F00";
        public string _SCH_WORD;

        public delegate void SendDataHandler(DataRow row);
        public event SendDataHandler DataRowSendEvent;

        private void PP003F00_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            this.Icon = ComnEtcFunc.GetFavicon();

            TxtSearch.EditValue = _SCH_WORD;
            BtnRetr.PerformClick();
            TxtSearch.Focus();
            this.ActiveControl = TxtSearch;
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            try
            {

                Cursor = Cursors.WaitCursor;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST1");
                dicParams.Add("FIND_IDX", CboSearch.SelectedIndex.ToString());
                dicParams.Add("FIND_WORD", TxtSearch.EditValue?.ToString());
                //dicParams.Add("USE_GB", RdoGubun.EditValue?.ToString());
                dicParams.Add("EXS_ID", FmMainToolBar2.drUser["USRCD"]?.ToString());

                if (GridViewRetr.RowCount > 0)
                {
                    GridRetr.DataSource = null;
                }

                GridRetr.DataSource = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                Cursor = Cursors.Default;

                if (GridViewRetr.RowCount == 0)
                {
                    TxtSearch.SelectAll();
                    TxtSearch.Focus();
                }
                else
                {
                    GridRetr.Focus();
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show("[BtnRetr Error] " + ex.Message);
            }
        }

        private void Bt_Add_Click(object sender, EventArgs e)
        {
            PP003F01 frm = new PP003F01();
            frm.Owner = this;
            frm.DataRowSendEvent += new PP003F01.SendDataHandler(GridRefresh);
            //frm.DataRowSendEvent += new
            frm.Show();
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            if (GridViewRetr.RowCount == 0)
            {
                XtraMessageBox.Show("사용자리스트의 내용이 존재하지 않습니다.");
                return;
            }

            RptApplSystemP1 pFrm = (RptApplSystemP1) this.Owner;
            if (pFrm.LineINFO(GridViewRetr.GetFocusedDataRow()))
            {
                DialogResult = DialogResult.OK;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void GridRefresh(string row)
        {
            BtnRetr.PerformClick();
            GridViewRetr.Focus();
            GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColGNAME2, row);
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

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
                BtnApply.PerformClick();
        }

        private void PP003F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
            {
                BtnRetr.PerformClick();
            }
            else if (e.KeyCode == Keys.F1)
            {
                Bt_Add.PerformClick();
            }
            else if(e.KeyCode == Keys.F4)
            {
                BtnDel.PerformClick();
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnApply.PerformClick();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void CboSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TxtSearch.Focus();
            }
        }

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr.Focus();
            }
        }

        private void BtnRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                BtnRetr.PerformClick();
            }
        }

        private void RdoGubun_EditValueChanged(object sender, EventArgs e)
        {
            BtnRetr.PerformClick();
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            try
            {
                string sPLNCD = GridViewRetr.GetFocusedRowCellValue("PLNCD")?.ToString();
                string sGNAME = GridViewRetr.GetFocusedRowCellValue("GNAME")?.ToString();

                if(string.IsNullOrEmpty(sPLNCD) || string.IsNullOrEmpty(sGNAME))
                {
                    XtraMessageBox.Show("삭제할 항목을 선택해주세요.");
                    return;
                }

                if (XtraMessageBox.Show("선택된 항목을 삭제하시겠습니까? \r\n삭제한 데이터는 복구할 수 없습니다."
                   , "결재라인 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "DEL1");
                dicParams.Add("PLNCD", sPLNCD);
                dicParams.Add("GNAME", sGNAME);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if(dt != null && dt.Rows.Count > 0)
                {
                    XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString());

                    if (dt.Rows[0]["RESULT"].Equals("1"))
                    {
                        int IDX = GridViewRetr.FocusedRowHandle;
                        BtnRetr.PerformClick();
                        GridViewRetr.FocusedRowHandle = IDX - 1;
                    }
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }
    }
}