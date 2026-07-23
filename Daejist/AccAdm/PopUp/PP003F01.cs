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
    public partial class PP003F01 : DevExpress.XtraEditors.XtraForm
    {
        private string PROCEDURE_ID = "DP_CM005F00";

        public delegate void SendDataHandler(string sVal);
        public event SendDataHandler DataRowSendEvent;
        private enum DataChanged { Changed, UnChanged }
        private DataChanged _UserChanged = DataChanged.UnChanged;
        private DataChanged _UserChanged2 = DataChanged.UnChanged;

        public PP003F01()
        {
            InitializeComponent();
        }

        private void PP003F01_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            this.Icon = ComnEtcFunc.GetFavicon();
            InitGrid();
        }

        #region [ 버튼 클릭 이벤트 ]
        // 초기화(F1)
        private void Bt_Reset_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();
            GridAdd.DataSource = dt;
            GridAdd2.DataSource = dt2;
        }

        // 저장(F3)
        private void Bt_Save_Click(object sender, EventArgs e)
        {
            string sGName = Tx_GName.EditValue?.ToString().Trim();
            string sMsg = string.Empty;

            if (string.IsNullOrEmpty(sGName))
            {
                XtraMessageBox.Show("결재선 명을 입력하세요.", "결재라인등록");
                return;
            }

            Dictionary<string, string> dicCheck = new Dictionary<string, string>();
            dicCheck.Clear();
            dicCheck.Add("CMD", "SGLine_Check");
            dicCheck.Add("GNAME", sGName);
            DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicCheck);
            if (dtChk.Rows.Count > 0)
            {
                XtraMessageBox.Show("해당 결재선 명이 이미 존재합니다. 결재선 명을 다시 입력해주세요.", "결재라인등록");
                return;
            }

            DataTable dt = (DataTable)GridAdd.DataSource;
            dt = DtUsrCdEmptyRemove(dt);

            DataTable dt2 = (DataTable)GridAdd2.DataSource;
            dt2 = DtUsrCdEmptyRemove(dt2);

            //if (dt != null)
            //{
            //    int i = 0;
            //    foreach (DataRow dr in dt.Rows)
            //    {
            //        if (string.IsNullOrEmpty(dt.Rows[i]["USRCD"].ToString()))
            //        {
            //            XtraMessageBox.Show("승인자를 선택하지 않은 행이 있습니다.");
            //            return;
            //        }
            //        i++;
            //    }
            //}
            //if (dt2 != null)
            //{
            //    foreach (DataRow dr in dt2.Rows)
            //    {
            //        int i = 0;
            //        if (string.IsNullOrEmpty(dt2.Rows[i]["USRCD"].ToString()))
            //        {
            //            XtraMessageBox.Show("참조자를 선택하지 않은 행이 있습니다.");
            //            return;
            //        }
            //        i++;
            //    }
            //}

            if (dt != null)
            {
                if (dt.Rows.Count == 0)
                {
                    XtraMessageBox.Show("결재자 정보가 없습니다.", "결재라인등록");
                    return;
                }
                try
                {
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sUsrCd = dr["USRCD"].ToString();
                        string sSeqNo = (i + 2).ToString();

                        Dictionary<string, string> dicParams = new Dictionary<string, string>();
                        dicParams.Clear();
                        dicParams.Add("CMD", "SGLine_Save");
                        dicParams.Add("PLNCD", FmMainToolBar2.drUser["USRCD"]?.ToString());
                        dicParams.Add("GNAME", sGName);
                        dicParams.Add("GUBUN", "0");    // 승인자 : 0, 참조자 : 1
                        dicParams.Add("SEQNO", sSeqNo);
                        dicParams.Add("APPRR", sUsrCd);

                        DataTable dtSave = DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
                        if (dtSave != null)
                        {
                            sMsg = dtSave.Rows[0]["MSG"].ToString();
                        }
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
                }
            }

            if (dt2 != null && dt2.Rows.Count > 0)
            {
                try
                {
                    int i = 0;
                    foreach (DataRow dr in dt2.Rows)
                    {
                        string sUsrCd = dr["USRCD"].ToString();
                        string sSeqNo = (i + 2).ToString();

                        Dictionary<string, string> dicParams = new Dictionary<string, string>();
                        dicParams.Clear();
                        dicParams.Add("CMD", "SGLine_Save");
                        dicParams.Add("PLNCD", FmMainToolBar2.drUser["USRCD"]?.ToString());
                        dicParams.Add("GNAME", sGName);
                        dicParams.Add("GUBUN", "1");    // 승인자 : 0, 참조자 : 1
                        dicParams.Add("SEQNO", sSeqNo);
                        dicParams.Add("APPRR", sUsrCd);

                        DataTable dtSave = DBConn.GetDataTable(DBConn.dbCon, PROCEDURE_ID, dicParams);
                        if(dtSave != null)
                        {
                            sMsg= dtSave.Rows[0]["MSG"].ToString();
                        }
                        i++;
                    }
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
                }
            }

            XtraMessageBox.Show(sMsg);
            DataRowSendEvent(sGName);
            Close();
        }

        private DataTable DtUsrCdEmptyRemove(DataTable dt)
        {
            DataTable result = dt.Clone();

            if (dt != null)
            {
                int cnt = 1;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];

                    string sUsrCd = row["USRCD"]?.ToString();

                    if (!string.IsNullOrEmpty(sUsrCd))
                    {
                        DataRow row_1 = result.NewRow();

                        row_1["SEQNO"] = cnt++;
                        row_1["USRCD"] = row["USRCD"];
                        row_1["USRNM"] = row["USRNM"];
                        row_1["USRID"] = row["USRID"];
                        //row_1["CHEKK"] = row["CHEKK"];

                        result.Rows.Add(row_1);
                    }
                }
            }

            return result;
        }

        // 닫기(ESC)
        private void Bt_Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        // 결재자 목록 버튼 (+/-)
        private void layoutControlGroup1_CustomButtonClick(object sender, DevExpress.XtraBars.Docking2010.BaseButtonEventArgs e)
        {
            if (e.Button.Properties.Tag.Equals("+"))
            {
                DataTable dt = (DataTable)GridAdd.DataSource;
                DataRow drnew = dt.NewRow();
                dt.Rows.Add(drnew);
                /// int i = 0 이 아니라 row의 갯수부터 시작해야함
                for (int i = dt.Rows.Count; i < dt.Rows.Count; ++i)
                {
                    dt.Rows[i]["SEQNO"] = i + 1;
                }
                GridAdd.DataSource = dt;
                GridViewAdd.FocusedRowHandle = dt.Rows.Count - 1;
                GridViewAdd.Focus();
                GridViewAdd.FocusedColumn = GridColUSRNM;
            }
            else if (e.Button.Properties.Tag.Equals("-"))
            {
                int iFocusedRow;
                DataTable dt = (DataTable)GridAdd.DataSource;
                if (dt.Rows.Count != 0)
                {
                    iFocusedRow = GridViewAdd.FocusedRowHandle;
                    dt.Rows.RemoveAt(iFocusedRow);
                }
                GridAdd.DataSource = dt;

                if (dt.Rows.Count != 0)
                {
                    GridViewAdd.FocusedRowHandle = dt.Rows.Count - 1;
                    GridViewAdd.Focus();
                    GridViewAdd.FocusedColumn = GridColUSRNM;
                }
            }
        }

        // 참조자 목록 버튼 (+/-)
        private void layoutControlGroup2_CustomButtonClick(object sender, DevExpress.XtraBars.Docking2010.BaseButtonEventArgs e)
        {
            if (e.Button.Properties.Tag.Equals("+"))
            {
                DataTable dt = (DataTable)GridAdd2.DataSource;
                DataRow drnew = dt.NewRow();
                dt.Rows.Add(drnew);
                /// int i = 0 이 아니라 row의 갯수부터 시작해야함
                for (int i = dt.Rows.Count; i < dt.Rows.Count; ++i)
                {
                    dt.Rows[i]["SEQNO"] = i + 1;
                }
                GridAdd2.DataSource = dt;
                GridViewAdd2.FocusedRowHandle = dt.Rows.Count - 1;
                GridViewAdd2.Focus();
                GridViewAdd2.FocusedColumn = GridColUSRNM2;
            }
            else if (e.Button.Properties.Tag.Equals("-"))
            {
                int iFocusedRow;
                DataTable dt = (DataTable)GridAdd2.DataSource;
                if (dt.Rows.Count != 0)
                {
                    iFocusedRow = GridViewAdd2.FocusedRowHandle;
                    dt.Rows.RemoveAt(iFocusedRow);
                }
                GridAdd2.DataSource = dt;

                if (dt.Rows.Count != 0)
                {
                    GridViewAdd2.FocusedRowHandle = dt.Rows.Count - 1;
                    GridViewAdd2.Focus();
                    GridViewAdd2.FocusedColumn = GridColUSRNM2;
                }
            }
        }
        #endregion

        #region [ 그리드 뷰 이벤트 ]
        // 승인자 버튼 클릭 시 팝업 로드 이벤트
        private void RepoUSRNM_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PP002F00 frm = new PP002F00();
            frm.Owner = this;
            frm.DataRowSendEvent += new PP002F00.SendDataHandler(USERINFO);
            frm.ShowDialog();

            _UserChanged = DataChanged.UnChanged;
        }

        // 승인자 값 변경 시 구분자 변경 이벤트
        private void RepoUSRNM_EditValueChanged(object sender, EventArgs e)
        {
            _UserChanged = DataChanged.Changed;
        }

        // 승인자 엔터키 입력 시 Row 자동생성
        private void RepoUSRNM_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //Focused Row Index가 마지막 행이 아닐 경우 리턴
                if (GridViewAdd.FocusedRowHandle != GridViewAdd.RowCount - 1)
                    return;

                // 마지막행의 데이터가 입력되어있지 않을 경우 해당 행의 승인자 컬럼 Focus
                if (string.IsNullOrEmpty(GridViewAdd.GetFocusedRowCellValue(GridColUSRNM)?.ToString()))
                {
                    GridViewAdd.FocusedColumn = GridColUSRNM;
                    return;
                }

                DataTable dt = (DataTable)GridAdd.DataSource;
                DataRow drnew = dt.NewRow();
                dt.Rows.Add(drnew);
                /// int i = 0 이 아니라 row의 갯수부터 시작해야함
                for (int i = dt.Rows.Count; i < dt.Rows.Count; ++i)
                {
                    dt.Rows[i]["SEQNO"] = i + 1;
                }
                GridAdd.DataSource = dt;
                GridViewAdd.FocusedRowHandle = dt.Rows.Count - 1;
                GridViewAdd.Focus();
                GridViewAdd.FocusedColumn = GridColUSRNM;
            }
        }

        // 승인자 떠날 시 이벤트
        private void RepoUSRNM_Leave(object sender, EventArgs e)
        {
            if (_UserChanged == DataChanged.UnChanged)
                return;

            ButtonEdit btnEdit = (ButtonEdit)sender;
            string sVal = btnEdit.EditValue?.ToString().Trim();
            if (string.IsNullOrEmpty(sVal))
                return;
            try
            {
                Cursor = Cursors.WaitCursor;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("CMD", "USRINFO");
                dicParams.Add("FIND_WORD", sVal);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                //결과행이 0일 경우 컨트롤에 바로 바인딩
                if (dt.Rows.Count == 1)
                {
                    USERINFO(dt.Rows[0]);
                }
                else
                {
                    PP002F00 frm = new PP002F00();
                    frm.Owner = this;
                    frm.DataRowSendEvent += new PP002F00.SendDataHandler(USERINFO);
                    frm._SCH_WORD = sVal;
                    frm.ShowDialog();
                }
                _UserChanged = DataChanged.UnChanged;
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        // 참조자 버튼 클릭 시 팝업 로드 이벤트
        private void RepoUSRNM2_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PP002F00 frm = new PP002F00();
            frm.Owner = this;
            frm.DataRowSendEvent += new PP002F00.SendDataHandler(USERINFO2);
            frm.ShowDialog();

            _UserChanged2 = DataChanged.UnChanged;
        }

        // 참조자 값 변경 시 구분자 변경 이벤트
        private void RepoUSRNM2_EditValueChanged(object sender, EventArgs e)
        {
            _UserChanged2 = DataChanged.Changed;
        }

        // 참조자 엔터키 입력 시 Row 자동생성
        private void RepoUSRNM2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //Focused Row Index가 마지막 행이 아닐 경우 리턴
                if (GridViewAdd2.FocusedRowHandle != GridViewAdd2.RowCount - 1)
                    return;

                // 마지막행의 데이터가 입력되어있지 않을 경우 해당 행의 승인자 컬럼 Focus
                if (string.IsNullOrEmpty(GridViewAdd2.GetFocusedRowCellValue(GridColUSRNM2)?.ToString()))
                {
                    GridViewAdd2.FocusedColumn = GridColUSRNM2;
                    return;
                }

                DataTable dt = (DataTable)GridAdd2.DataSource;
                DataRow drnew = dt.NewRow();
                dt.Rows.Add(drnew);
                /// int i = 0 이 아니라 row의 갯수부터 시작해야함
                for (int i = dt.Rows.Count; i < dt.Rows.Count; ++i)
                {
                    dt.Rows[i]["SEQNO"] = i + 1;
                }
                GridAdd2.DataSource = dt;
                GridViewAdd2.FocusedRowHandle = dt.Rows.Count - 1;
                GridViewAdd2.Focus();
                GridViewAdd2.FocusedColumn = GridColUSRNM2;
            }
        }

        // 참조자 떠날 시 이벤트
        private void RepoUSRNM2_Leave(object sender, EventArgs e)
        {
            if (_UserChanged2 == DataChanged.UnChanged)
                return;

            ButtonEdit btnEdit = (ButtonEdit)sender;
            string sVal = btnEdit.EditValue?.ToString().Trim();
            if (string.IsNullOrEmpty(sVal))
                return;
            try
            {
                Cursor = Cursors.WaitCursor;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Add("CMD", "USRINFO");
                dicParams.Add("FIND_WORD", sVal);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                //결과행이 0일 경우 컨트롤에 바로 바인딩
                if (dt.Rows.Count == 1)
                {
                    USERINFO2(dt.Rows[0]);
                }
                else
                {
                    PP002F00 frm = new PP002F00();
                    frm.Owner = this;
                    frm.DataRowSendEvent += new PP002F00.SendDataHandler(USERINFO2);
                    frm._SCH_WORD = sVal;
                    frm.ShowDialog();
                }
                _UserChanged2 = DataChanged.UnChanged;
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region [ 메서드 ]
        // 로드 시 그리드 초기화 메서드
        private void InitGrid()
        {
            DataTable dt = new DataTable();
            DataTable dt2 = new DataTable();

            dt.Columns.Add("USRCD");
            dt.Columns.Add("SEQNO");
            dt.Columns.Add("USRNM");
            dt.Columns.Add("USRID");
            GridAdd.DataSource = dt;

            dt2.Columns.Add("USRCD");
            dt2.Columns.Add("SEQNO");
            dt2.Columns.Add("USRNM");
            dt2.Columns.Add("USRID");
            GridAdd2.DataSource = dt2;
        }

        // 승인자 정보 자동 바인딩 메서드
        public void USERINFO(DataRow row)
        {
            string sUSRCD1 = row["USRCD"].ToString();

            DataTable dt = (DataTable)GridAdd.DataSource;
            for (int i = 0; i < dt.Rows.Count; ++i)
            {
                string sUSRCD2 = dt.Rows[i]["USRCD"].ToString();

                if (sUSRCD2 == sUSRCD1)
                {
                    MessageBox.Show("결재자 중복입니다 다시 선택해 주세요");
                    return;
                }
            }
            GridViewAdd.SetFocusedRowCellValue(GridColUSRNM, row["USRNM"]);
            GridViewAdd.SetFocusedRowCellValue(GridColUSRCD, row["USRCD"]);
            GridViewAdd.SetFocusedRowCellValue(GridColUSRID, row["USRID"]);
        }

        // 참조자 정보 자동 바인딩 메서드
        public void USERINFO2(DataRow row)
        {
            string sUSRCD1 = row["USRCD"].ToString();

            DataTable dt = (DataTable)GridAdd2.DataSource;
            for (int i = 0; i < dt.Rows.Count; ++i)
            {
                string sUSRCD2 = dt.Rows[i]["USRCD"].ToString();

                if (sUSRCD2 == sUSRCD1)
                {
                    MessageBox.Show("참조자 중복입니다 다시 선택해 주세요");
                    return;
                }
            }
            GridViewAdd2.SetFocusedRowCellValue(GridColUSRNM2, row["USRNM"]);
            GridViewAdd2.SetFocusedRowCellValue(GridColUSRCD2, row["USRCD"]);
            GridViewAdd2.SetFocusedRowCellValue(GridColUSRID2, row["USRID"]);
        }
        #endregion

        #region [ 키 이벤트 ]
        private void PP003F01_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                Bt_Reset.PerformClick();
            }
            else if (e.KeyCode == Keys.F3)
            {
                Bt_Save.PerformClick();
            }
            else if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
        #endregion

        #region [ Row 스타일 ]
        private void GridViewAdd_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewAdd_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewAdd2_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewAdd2_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }
        #endregion
    }
}