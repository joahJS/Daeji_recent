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
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Data.SqlClient;

using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace AccAdm
{
    public partial class RptApplSystemP1 : DevExpress.XtraEditors.XtraForm
    {
        public RptApplSystemP1()
        {
            InitializeComponent();
        }

        //FTP 서버
        private string sInitDir = ComnEtcFunc.FTP_ROOT + @"/ERP/AprlDocs/"; //양식
        private string sInitDir2 = ComnEtcFunc.FTP_ROOT + @"/ERP/AttFile/"; //첨부파일
        private string user = ComnEtcFunc.FTP_USER;
        private string pw = ComnEtcFunc.FTP_PW;

        //임시 파일저장 경로
        private string tempDoctPath = Application.StartupPath + @"/tempDoct/"+FmMainToolBar2.drUser["USRCD"];
        //임시 파일명
        private string _tempFileName = "전자결재 문서 작성.xlsx";

        private string PROCEDURE_ID = "DP_SI003F01";
        private string _FILNM;

        public string _SLINO;
        public string _DOCTP;
        public string _CUSTOM;
        public decimal _PAY;
        public string AddModifyGb { get; set; }

        public delegate void SendDataHandler(Dictionary<string, object> dicParams);
        public event SendDataHandler DataRowSendEvent;

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        private void RptApplSystemP1_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            ComnEtcFunc.gp_SetColorFocused(layoutControl3);

            DataTable dtDept = GetLookupData("0");
            ComGrid.SetLookUpEdit(LkupDept, dtDept, "CD", "NM", "");
            DataTable dtUsr = GetLookupData("1");
            ComGrid.SetLookUpEdit(LkupUsr, dtUsr, "CD", "NM", "");
            DataTable dtDoctp = GetLookupData("2");
            dtDoctp.Columns["SEQ"].ColumnMapping = MappingType.Hidden; // SEQ 컬럼 안보이게 설정
            ComGrid.SetLookUpEdit(LkupDOCTP, dtDoctp, "CD", "NM", "");

            SetValue();
            //GetUploadFileInfo();
            //ExcelCloseAndDelete();
        }

        private void SetValue()
        {
            try
            {
                GridRetr2.DataSource = SetLoadPackage();
                GridRetr3.DataSource = SetLoadPackage3();

                //경영관리부만 문서잠금 체크
                string sDept = FmMainToolBar2.drUser["DEPTCD"]?.ToString();
                if(!string.IsNullOrEmpty(sDept) && sDept.Equals("2000"))
                {
                    layoutControlItem22.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }

                if (AddModifyGb.Equals("MOD"))
                {
                    Dictionary<string, string> dicParams = new Dictionary<string, string>();
                    dicParams.Clear();
                    dicParams.Add("CMD", "MSTR_RETR");
                    dicParams.Add("SLINO", _SLINO);

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                    string sState = string.Empty;
                    string sMSG = string.Empty;

                    if (dt.Rows.Count != 0)
                    {
                        TxtSLINO.EditValue = _SLINO;  //문서번호
                        TDate.EditValue = dt.Rows[0]["TDATE"];// 작성일
                        TxtSTITL.EditValue = dt.Rows[0]["STITL"];//제목
                        LkupDOCTP.EditValue = dt.Rows[0]["DOCTP"];//문서구분
                        _FILNM = dt.Rows[0]["FILNM"]?.ToString();//파일명
                        LkupUsr.EditValue = dt.Rows[0]["PLNCD"];//작성자코드
                        LkupDept.EditValue = dt.Rows[0]["DEPTCD"]; // 부서
                        ChkDlock.EditValue = dt.Rows[0]["DLOCK"];//문서잠금

                        sState = dt.Rows[0]["STATE"]?.ToString();

                        if (!dt.Rows[0]["PLNCD"].Equals(FmMainToolBar2.drUser["USRCD"]))
                        {
                            SetEnable();
                            sMSG = "본인글 외에는 조회만 가능합니다.\r\n";
                        }
                        else if (sState.Equals("2"))//진행중
                        {
                            sMSG += "해당건은 결재 진행중이므로 수정이 불가능 합니다.";
                            SetEnable();
                        }
                        //반려 수정가능하도록 변경(업체요청) 2022-12-07
                        //else if (sState.Equals("3"))//반려
                        //{
                        //    sMSG += "해당건은 결재 반려되었으므로 수정이 불가능 합니다.";
                        //    SetEnable();
                        //}
                        else if (sState.Equals("4"))//완료
                        {
                            sMSG += "해당건은 결재 완료되었으므로 수정이 불가능 합니다.";
                            SetEnable();
                        }
                        //최종반려인 경우 수정 불가
                        else if (sState.Equals("5"))//최종반려
                        {
                            sMSG += "해당건은 최종반려되었으므로 수정이 불가능 합니다.";
                            SetEnable();
                        }

                        //파일다운로드
                        if (!DBNull.Value.Equals(dt.Rows[0]["EXCFIL"]))
                        {
                            string downPath = tempDoctPath + @"/"+ _tempFileName;
                            byte[] file = (byte[])dt.Rows[0]["EXCFIL"];

                            if (!Directory.Exists(tempDoctPath))
                            {
                                Directory.CreateDirectory(tempDoctPath);
                            }

                            if (file != null)
                            {
                                FileStream fs;

                                fs = new FileStream(downPath, FileMode.OpenOrCreate, FileAccess.Write);
                                fs.Write(file, 0, file.Length);
                                fs.Close();
                            }
                        }
                    }

                    //결재선
                    dicParams.Clear();
                    dicParams.Add("CMD", "LINE_RETR1");
                    dicParams.Add("SLINO", _SLINO);
                    DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                    if (dt.Rows.Count != 0)
                    {
                        GridRetr2.DataSource = dt2;
                    }

                    //참조
                    dicParams.Clear();
                    dicParams.Add("CMD", "LINE_RETR2");
                    dicParams.Add("SLINO", _SLINO);
                    DataTable dt3 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);
                    if (dt.Rows.Count != 0)
                    {
                        GridRetr3.DataSource = dt3;
                    }

                    label1.Text = sMSG;
                    layoutControlItem21.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }
                else if (AddModifyGb.Equals("ADD"))
                {
                    TDate.EditValue = DateTime.Now;// 작성일

                    LkupUsr.EditValue = FmMainToolBar2.drUser["USRCD"];//작성자코드
                    LkupDept.EditValue = FmMainToolBar2.drUser["DEPTCD"]; // 작성자부서

                    layoutControlItem12.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                }
                else if (AddModifyGb.Equals("COPY"))
                {
                    TDate.EditValue = DateTime.Now;// 작성일

                    LkupUsr.EditValue = FmMainToolBar2.drUser["USRCD"];//작성자코드
                    LkupDept.EditValue = FmMainToolBar2.drUser["DEPTCD"]; // 작성자부서
                    LkupDOCTP.EditValue = _DOCTP;

                    layoutControlItem13.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;

                    ExcelCloseAndDelete();
                    string sDownPath = EditDoct();

                    if (!string.IsNullOrEmpty(sDownPath))
                    {
                        EditAndSaveExcelFile(sDownPath);
                        //Process.Start(sDownPath);
                    }
                    //BtnDoctCopy.PerformClick();
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void SetEnable()
        {
            BtnSave.Enabled = false;
            //BtnUpload.Enabled = false;
            //bt_LineAdd.Enabled = false;
            //BtnWritDoct.Enabled = false;
            //BtnNewWrit.Enabled = false;
            //BtnDoctCopy.Enabled = false;

            layoutControlGroup5.CustomButtonClick -= new DevExpress.XtraBars.Docking2010.BaseButtonEventHandler(this.layoutControlGroup5_CustomButtonClick);
            layoutControlGroup6.CustomButtonClick -= new DevExpress.XtraBars.Docking2010.BaseButtonEventHandler(this.layoutControlGroup6_CustomButtonClick);
        }

        #region [결재라인 관련]
        private enum DataChanged { Changed, UnChanged }
        private DataChanged _UserChanged1 = DataChanged.UnChanged;
        private DataChanged _UserChanged2 = DataChanged.UnChanged;
        private DataChanged _UserChanged3 = DataChanged.UnChanged;
        private DataChanged _UserChanged4 = DataChanged.UnChanged;
        private DataChanged _LineChange = DataChanged.UnChanged;
        private DataChanged _PJCODEChange = DataChanged.UnChanged;

        private DataTable SetLoadPackage()
        {
            string sUSRNM = FmMainToolBar2.drUser["USRNM"]?.ToString();
            string sUSRCD = FmMainToolBar2.drUser["USRCD"]?.ToString();
            string sUSRID = FmMainToolBar2.drUser["USRID"]?.ToString();

            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine("  DECLARE @SEQN NUMERIC(2)=1;               ");
            strSql.AppendLine(" SELECT @SEQN AS SEQNO                      ");
            strSql.AppendLine(" 	 ,'" + sUSRNM + "' AS USRNM             ");
            strSql.AppendLine(" 	 ,'" + sUSRID + "' AS USRID           ");
            strSql.AppendLine(" 	 ,CONVERT(NUMERIC," + sUSRCD + ") AS USRCD            ");
            strSql.AppendLine(" 	 , 'N' AS CHEKK                        ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            return dt;
        }

        private DataTable SetLoadPackage3()
        {
            string sUSRNM = FmMainToolBar2.drUser["USRNM"]?.ToString();
            string sUSRCD = FmMainToolBar2.drUser["USRCD"]?.ToString();
            string sUSRID = FmMainToolBar2.drUser["USRID"]?.ToString();

            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine("  DECLARE @SEQN NUMERIC(2)=1;             ");
            strSql.AppendLine(" SELECT @SEQN AS SEQNO                    ");
            strSql.AppendLine(" 	 ,'' AS USRNM                        ");
            strSql.AppendLine(" 	 ,'' AS USRID                        ");
            strSql.AppendLine(" 	 ,NULL AS USRCD                        ");
            strSql.AppendLine(" 	 , 'N' AS CHEKK                      ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            return dt;
        }

        #region [결재라인 불러오기]
        private void bt_LineAdd_Click(object sender, EventArgs e)
        {
            PP003F00 frm = new PP003F00();
            frm.Owner = this;
            //frm.DataRowSendEvent += new PP003F00.SendDataHandler(LineINFO);
            frm.ShowDialog();
            _LineChange = DataChanged.UnChanged;
        }
        public bool LineINFO(DataRow row)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                Dictionary<string, string> dicParams = new Dictionary<string, string>();
                /////////////////////숭인자
                dicParams.Clear();
                dicParams.Add("CMD", "LIST3");
                dicParams.Add("PLNCD", row["PLNCD"].ToString());
                dicParams.Add("GNAME", row["GNAME"].ToString());

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                string sUSRCD1 = FmMainToolBar2.drUser["USRCD"]?.ToString();

                for (int i = 0; i < dt.Rows.Count; ++i)
                {
                    string sUSRCD2 = dt.Rows[i]["USRCD"].ToString();

                    if (sUSRCD2.Equals(sUSRCD1))
                    {
                        MessageBox.Show("결재 라인에 등록자가 있어 중복됩니다 다시 선택해 주세요");

                        Cursor = Cursors.Default;
                        return false;
                    }

                }

                DataTable dtt = SetLoadPackage();
                dtt.Merge(dt);
                GridRetr2.DataSource = dtt;

                ///////////////////////참조자
                dicParams.Clear();
                dicParams.Add("CMD", "LIST4");
                dicParams.Add("PLNCD", row["PLNCD"].ToString());
                dicParams.Add("GNAME", row["GNAME"].ToString());

                DataTable dt2 = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                GridRetr3.DataSource = dt2;
                Cursor = Cursors.Default;

                return true;

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                Cursor = Cursors.Default;
                return false;
            }
        }
        #endregion

        #region [결재자 추가]
        private void layoutControlGroup5_CustomButtonClick(object sender, DevExpress.XtraBars.Docking2010.BaseButtonEventArgs e)
        {
            ComnGridFunc.GridViewLineAdd_Click(GridViewRetr2, GridColSEQNO2, GridColUSRNM);
        }
        #endregion

        #region [결재자 삭제]
        private void RepoDelete_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            DataRow row = GridViewRetr2.GetFocusedDataRow();

            if (GridViewRetr2.RowCount == 1)
            {
                string sDsg = "첫번째 행은 삭제할 수 없습니다";
                MessageBox.Show(sDsg);
                return;
            }
            DataTable dt = (DataTable)GridRetr2.DataSource;
            dt.Rows.Remove(row);
            for (int i = 0; i < dt.Rows.Count; ++i)
            {
                dt.Rows[i]["SEQNO"] = i + 1;
            }
            GridRetr2.DataSource = dt;
        }
        #endregion

        #region [결재자 선택]
        private void RepoBtnUser_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PP002F00 frm = new PP002F00();
            frm.Owner = this;
            frm.DataRowSendEvent += new PP002F00.SendDataHandler(USERINFO3);
            frm._SCH_WORD = GridViewRetr2.GetFocusedRowCellValue(GridColUSRNM)?.ToString();
            frm.ShowDialog();

            _UserChanged3 = DataChanged.UnChanged;
        }
        private void RepoBtnUser_EditValueChanged(object sender, EventArgs e)
        {
            _UserChanged3 = DataChanged.Changed;
        }
        public void USERINFO3(DataRow row)
        {
            string sUSRCD1 = row["USRCD"].ToString();

            DataTable dt = (DataTable)GridRetr2.DataSource;
            for (int i = 0; i < dt.Rows.Count; ++i)
            {
                string sUSRCD2 = dt.Rows[i]["USRCD"].ToString();

                if (sUSRCD2 == sUSRCD1)
                {
                    MessageBox.Show("결재자 중복입니다 다시 선택해 주세요");
                    return;
                }

            }
            GridViewRetr2.SetFocusedRowCellValue(GridColUSRNM, row["USRNM"]);
            GridViewRetr2.SetFocusedRowCellValue(GridColUSRCD, row["USRCD"]);
            GridViewRetr2.SetFocusedRowCellValue("USRID", row["USRID"]);
        }
        private void RepoBtnUser_Leave(object sender, EventArgs e)
        {
            if (_UserChanged3 == DataChanged.UnChanged)
                return;

            ButtonEdit btnEdit = (ButtonEdit)sender;
            string sVal = btnEdit.EditValue?.ToString().Trim();
            if (string.IsNullOrEmpty(sVal))
            {

                return;
            }
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
                    USERINFO3(dt.Rows[0]);
                }
                else
                {
                    PP002F00 frm = new PP002F00();
                    frm.Owner = this;
                    frm.DataRowSendEvent += new PP002F00.SendDataHandler(USERINFO3);
                    frm._SCH_WORD = sVal;
                    frm.ShowDialog();
                }
                _UserChanged3 = DataChanged.UnChanged;
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void RepoBtnUser_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //Focused Row Index가 마지막 행이 아닐 경우 리턴
                if (GridViewRetr2.FocusedRowHandle != GridViewRetr2.RowCount - 1)
                    return;

                //마지막행의 사람이 입력되어있지 않을 경우 추가하지 말고 해당 행의 품명으로 입력받게끔 Focus처리
                if (string.IsNullOrEmpty(GridViewRetr2.GetFocusedRowCellValue(GridColUSRNM)?.ToString()))
                {
                    GridViewRetr2.FocusedColumn = GridColUSRNM;
                    return;
                }

                DataTable dtFix = (DataTable)GridRetr2.DataSource;
                DataRow drNew = dtFix.NewRow();
                int lastnum = GridViewRetr2.RowCount - 1;
                drNew["SEQNO"] = Convert.ToInt16(dtFix.Rows[lastnum]["SEQNO"]?.ToString()) + 1;
                drNew["CHEKK"] = "N";

                dtFix.Rows.Add(drNew);

                GridRetr2.DataSource = dtFix;
                GridViewRetr2.FocusedRowHandle = dtFix.Rows.Count - 1;
                GridViewRetr2.FocusedColumn = GridColUSRNM;

            }

        }
        #endregion

        #region [참조자 선택]
        public void USERINFO4(DataRow row)
        {
            GridViewRetr3.SetFocusedRowCellValue(GridColUSRNM2, row["USRNM"]);
            GridViewRetr3.SetFocusedRowCellValue(GridColUSRCD2, row["USRCD"]);
            GridViewRetr3.SetFocusedRowCellValue("USRID", row["USRID"]);
        }
        private void RepoBtnUser2_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PP002F00 frm = new PP002F00();
            frm.Owner = this;
            frm.DataRowSendEvent += new PP002F00.SendDataHandler(USERINFO4);
            frm._SCH_WORD = GridViewRetr3.GetFocusedRowCellValue(GridColUSRNM)?.ToString();
            frm.ShowDialog();
            _UserChanged4 = DataChanged.UnChanged;
        }
        private void RepoBtnUser2_EditValueChanged(object sender, EventArgs e)
        {
            _UserChanged4 = DataChanged.Changed;
        }
        private void RepoBtnUser2_Leave(object sender, EventArgs e)
        {
            if (_UserChanged4 == DataChanged.UnChanged)
                return;

            ButtonEdit btnEdit = (ButtonEdit)sender;
            string sVal = btnEdit.EditValue?.ToString().Trim();
            if (string.IsNullOrEmpty(sVal))
            {
                return;
            }
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
                    USERINFO4(dt.Rows[0]);
                }
                else
                {
                    PP002F00 frm = new PP002F00();
                    frm.Owner = this;
                    frm.DataRowSendEvent += new PP002F00.SendDataHandler(USERINFO4);
                    frm._SCH_WORD = sVal;
                    frm.ShowDialog();
                }
                _UserChanged4 = DataChanged.UnChanged;
                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }
        private void RepoBtnUser2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //Focused Row Index가 마지막 행이 아닐 경우 리턴
                if (GridViewRetr3.FocusedRowHandle != GridViewRetr3.RowCount - 1)
                    return;

                //마지막행의 사람이 입력되어있지 않을 경우 추가하지 말고 해당 행의 품명으로 입력받게끔 Focus처리
                if (string.IsNullOrEmpty(GridViewRetr3.GetFocusedRowCellValue(GridColUSRNM2)?.ToString()))
                {
                    GridViewRetr3.FocusedColumn = GridColUSRNM2;
                    return;
                }
                DataTable dtFix = (DataTable)GridRetr3.DataSource;
                DataRow drNew = dtFix.NewRow();
                int lastnum = GridViewRetr3.RowCount - 1;
                drNew["SEQNO"] = Convert.ToInt16(dtFix.Rows[lastnum]["SEQNO"]?.ToString()) + 1;

                dtFix.Rows.Add(drNew);
                GridRetr3.DataSource = dtFix;
                GridViewRetr3.FocusedRowHandle = dtFix.Rows.Count - 1;
                GridViewRetr3.FocusedColumn = GridColUSRNM;
            }
        }
        #endregion

        #region [참조자 추가]
        private void layoutControlGroup6_CustomButtonClick(object sender, DevExpress.XtraBars.Docking2010.BaseButtonEventArgs e)
        {
            ComnGridFunc.GridViewLineAdd_Click(GridViewRetr3, GridColSEQNO3, GridColUSRNM2);
        }
        #endregion

        #region [참조자 삭제]
        private void RepoDelete2_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            if (GridViewRetr3.FocusedRowHandle == 0)
            {
                GridViewRetr3.SetFocusedRowCellValue(GridColUSRNM2, string.Empty);
                GridViewRetr3.SetFocusedRowCellValue(GridColUSRCD2, string.Empty);
                GridViewRetr3.SetFocusedRowCellValue("USRID", string.Empty);
            }
            ComnGridFunc.GridViewLineDelete_Click(GridViewRetr3, GridRetr3, GridColUSRNM2, "SEQNO");
        }
        #endregion

        #endregion

        #region LOOKUP DATA
        private DataTable GetLookupData(string sGB)
        {
            StringBuilder strSql = new StringBuilder();

            if (sGB.Equals("0"))
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT DEPT_CD AS CD");
                strSql.AppendLine("      , DEPT_NM AS NM");
                strSql.AppendLine("   FROM ACC_DEPT_CD  ");
                strSql.AppendLine("  WHERE USE_YN = 'Y' ");
            }
            else if (sGB.Equals("1"))
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT USRCD AS CD ");
                strSql.AppendLine("      , USRNM AS NM ");
                strSql.AppendLine("   FROM ZUSRLST     ");
                strSql.AppendLine("  WHERE USEYN = 'Y' ");
            }
            else if (sGB.Equals("2"))
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT DOCTP AS CD                                  ");
                strSql.AppendLine("      , DOCNM AS NM                                  ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY LEN(DOCTP), DOCTP) AS SEQ");
                strSql.AppendLine("   FROM DOCT_K                                       ");
            }

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }
        #endregion

        private void RptApplSystemP1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            //DialogResult = DialogResult.Cancel;
            Dispose();
        }

        private DataTable DtUsrCdEmptyRemove(DataTable dt)
        {
            DataTable result = dt.Clone();

            if(dt != null)
            {
                int cnt = 1;
                for(int i = 0; i < dt.Rows.Count; i++)
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

        private void BtnSave_Click(object sender, EventArgs e)
        {

            Dictionary<string, object> dicParams = new Dictionary<string, object>();

            try
            {
                Process[] prcList = Process.GetProcessesByName("Excel");
                if (prcList.Length != 0)
                {
                    if (XtraMessageBox.Show("작성중인 결재 문서를 종료하시겠습니까? \r\n저장하지 않은 데이터는 복구할 수 없습니다.", "엑셀문서 확인", MessageBoxButtons.YesNo) != DialogResult.Yes)
                    {
                        return;
                    }

                    for (int i = 0; i < prcList.Length; i++)
                    {
                        Process prcessInfo = prcList[i];

                        string sProcessFileName = prcessInfo.MainWindowTitle;

                        if (sProcessFileName.Equals(_tempFileName + " - Excel"))
                        {
                            int processId = prcessInfo.Id;

                            if (processId != 0)
                            {
                                System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById(processId);
                                excelProcess.CloseMainWindow();
                                excelProcess.Refresh();
                                excelProcess.Kill();
                            }
                        }
                    }

                    Thread.Sleep(1000);
                }

                if (AddModifyGb.Equals("ADD") || AddModifyGb.Equals("COPY"))
                {
                    string deptnm = LkupDept.Properties.GetDisplayText(LkupDept.EditValue);

                    if (deptnm == null || deptnm.Equals(""))
                    {
                        XtraMessageBox.Show("부서를 선택해주세요.");
                        LkupDept.Focus();
                        return;
                    }

                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    strSql.AppendLine(" SELECT '" + deptnm.Substring(0, 2) + " 제' + CONVERT(VARCHAR(8), GETDATE(), 23) + FORMAT((ISNULL(MAX(CONVERT(NUMERIC,SUBSTRING(SLINO,13,3))),0) + 1), '000') AS SLINO ");
                    strSql.AppendLine("   FROM DOCT_M A1                                                                              ");
                    strSql.AppendLine("  WHERE SLINO LIKE '" + deptnm.Substring(0, 2) + " 제' + CONVERT(VARCHAR(8), GETDATE(), 23) + '%'                        ");
                    
                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    _SLINO = dt.Rows[0]["SLINO"]?.ToString();
                    TxtSLINO.EditValue = _SLINO;
                    

                }

                string sTDATE = TDate.DateTime.ToString("yyyy-MM-dd");
                string sPLNCD = LkupUsr.EditValue?.ToString();
                string sSTITL = TxtSTITL.EditValue?.ToString();
                string sDOCTP = LkupDOCTP.EditValue?.ToString();
                //string sFILNM = TxtFileNm.EditValue?.ToString();
                string sDeptcd = LkupDept.EditValue?.ToString();
                string sDlock = ChkDlock.EditValue?.ToString();
                string sUSER = FmMainToolBar2.UserID;

                //2023.11.2일  거래처명,금액 DB 저장
                if (sDOCTP == "23"  )  //지출결의서 문서만 실행
                {
                    SerchExcelFile(tempDoctPath + @"/" + _tempFileName, 23);
                }
                if (sDOCTP == "37")  //계약서 문서만 실행
                {
                    SerchExcelFile(tempDoctPath + @"/" + _tempFileName, 37);
                }

                dicParams.Add("CMD", "SAVE");
                dicParams.Add("SLINO", _SLINO);     // 문서번호
                dicParams.Add("TDATE", sTDATE);     // 작성일
                dicParams.Add("PLNCD", sPLNCD);     // 작성자코드
                dicParams.Add("DEPTCD", sDeptcd);   //부서
                dicParams.Add("STITL", sSTITL);     // 제목
                dicParams.Add("DOCTP", sDOCTP);     // 문서구분
                dicParams.Add("DLOCK", sDlock);     //문서잠금
                dicParams.Add("CUSTOM", _CUSTOM);   //거래처명
                dicParams.Add("PAY", _PAY);         //금액

                byte[] bFile = null;

                FileInfo fileInfo = new FileInfo(tempDoctPath + @"/" + _tempFileName);
                if (fileInfo.Exists)
                {
                    EditAndSaveExcelFile(tempDoctPath + @"/" + _tempFileName);
                    bFile = File.ReadAllBytes(tempDoctPath + @"/" + _tempFileName);
                }

                if (bFile != null)
                {
                    dicParams.Add("EXCFIL", bFile);
                    //dicParams.Add("FILNM", sFILNM + sExtFileNm);
                }
                dicParams.Add("SUSER", sUSER);

                DataTable dt2 = (DataTable)GridRetr2.DataSource;
                dt2 = DtUsrCdEmptyRemove(dt2);
                if (dt2 != null && dt2.Rows.Count > 0)
                {
                    string json2 = ComnEtcFunc.DataTableToJsonObj(dt2);
                    dicParams.Add("JSON2", json2);
                }

                DataTable dt3 = (DataTable)GridRetr3.DataSource;
                dt3 = DtUsrCdEmptyRemove(dt3);
                if (dt3 != null && dt3.Rows.Count > 0)
                {
                    string json3 = ComnEtcFunc.DataTableToJsonObj(dt3);
                    dicParams.Add("JSON3", json3);
                }

                DataTable resultDt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);



                if (resultDt != null)
                {
                    if (resultDt.Rows[0]["RESULT"].Equals("1"))
                    {
                        #region 첨부파일 저장기능 제거 요청 2022-07-19
                        //첨부파일 저장
                        //DataTable dtAtt = GridRetr1.DataSource as DataTable;

                        //DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dtAtt);
                        //DataTable dtMerge = dsSave.Tables[0];
                        //DataTable dtDel = dtAtt.GetChanges(DataRowState.Deleted);

                        //StringBuilder strSql = new StringBuilder();

                        //string AttchPath = sInitDir2 + "/" + _SLINO;

                        //try
                        //{
                        //    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                        //    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                        //    cmd.Transaction = DBConn.dbTran;

                        //    if (dtDel != null && dtDel.Rows.Count > 0)//삭제
                        //    {
                        //        for(int i = 0; i < dtDel.Rows.Count; i++)
                        //        {
                        //            string sSeqno = dtDel.Rows[i]["SEQNO", DataRowVersion.Original]?.ToString();
                        //            string sFilnm = dtDel.Rows[i]["FILNM", DataRowVersion.Original]?.ToString();

                        //            strSql.Clear();
                        //            strSql.AppendLine(" DELETE FROM DOCT_F");
                        //            strSql.AppendLine("  WHERE SLINO = '"+ _SLINO + "' ");
                        //            strSql.AppendLine("    AND SEQNO ="+ sSeqno + " ");

                        //            cmd.CommandType = CommandType.Text;
                        //            cmd.CommandText = strSql.ToString();
                        //            cmd.ExecuteNonQuery();

                        //            if (!string.IsNullOrEmpty(sFilnm))
                        //            {
                        //                ComnEtcFunc.DeleteFTPFile(AttchPath + "/" + sFilnm, user, pw);
                        //            }
                        //        }
                        //    }

                        //    if(dtMerge != null && dtMerge.Rows.Count > 0)
                        //    {
                        //        for (int i=0;i< dtMerge.Rows.Count; i++)
                        //        {
                        //            string sSeqno = dtMerge.Rows[i]["SEQNO"]?.ToString();
                        //            string sFilnm = dtMerge.Rows[i]["FILNM"]?.ToString();
                        //            string sOPATH = dtMerge.Rows[i]["OPATH"]?.ToString();

                        //            byte[] FILE = File.ReadAllBytes(sOPATH);

                        //            strSql.Clear();
                        //            if (string.IsNullOrEmpty(sSeqno)) //신규추가
                        //            {
                        //                strSql.AppendLine("         DECLARE @SEQNO NUMERIC;                                              ");
                        //                strSql.AppendLine("                                                                                  ");
                        //                strSql.AppendLine("          SELECT @SEQNO = ISNULL(MAX(SEQNO), 0) + 1");
                        //                strSql.AppendLine("            FROM DOCT_F                                                           ");
                        //                strSql.AppendLine("           WHERE SLINO = '" + _SLINO + "'                                                       ");
                        //                strSql.AppendLine("          INSERT INTO DOCT_F( SLINO                                               ");
                        //                strSql.AppendLine("                            , SEQNO                                              ");
                        //                strSql.AppendLine("                            , FILNM                                              ");
                        //                strSql.AppendLine("                            , CUSER                                              ");
                        //                strSql.AppendLine("                            , CDATE)                                             ");
                        //                strSql.AppendLine("                                                                                 ");
                        //                strSql.AppendLine("                      VALUES('" + _SLINO + "'                                                  ");
                        //                strSql.AppendLine("                            , @SEQNO                                             ");
                        //                strSql.AppendLine("                            , '" + sFilnm + "'                                                 ");
                        //                strSql.AppendLine("                            , '" + sUSER + "'                                                 ");
                        //                strSql.AppendLine("                            , CONVERT(VARCHAR(20), GETDATE(), 20))               ");
                        //            }
                        //            else
                        //            {
                        //                strSql.AppendLine(" IF EXISTS(SELECT* FROM DOCT_F WHERE SLINO = '"+ _SLINO + "' AND SEQNO = "+ sSeqno + ") ");
                        //                strSql.AppendLine("    BEGIN                                                                         ");
                        //                strSql.AppendLine("          UPDATE DOCT_F                                                           ");
                        //                strSql.AppendLine("             SET FILNM = '"+ sFilnm + "'                                                       ");
                        //             strSql.AppendLine("               , MUSER = '"+ sUSER + "'                                                       ");
                        //             strSql.AppendLine("               , MDATE = CONVERT(VARCHAR(20), GETDATE(), 20)                      ");
                        //                strSql.AppendLine("           WHERE SLINO = '" + _SLINO + "'                                        ");
                        //                strSql.AppendLine("             AND SEQNO = "+ sSeqno                                          );
                        //                strSql.AppendLine("      END                                                                         ");
                        //                strSql.AppendLine(" ELSE                                                                             ");
                        //                strSql.AppendLine("    BEGIN                                                                         ");
                        //                strSql.AppendLine("         DECLARE @SEQNO NUMERIC;                                              ");
                        //                strSql.AppendLine("                                                                                  ");
                        //                strSql.AppendLine("          SELECT @SEQNO = ISNULL(MAX(SEQNO), 0) + 1");
                        //                strSql.AppendLine("            FROM DOCT_F                                                           ");
                        //                strSql.AppendLine("           WHERE SLINO = '" + _SLINO + "'                                                       ");
                        //                strSql.AppendLine("          INSERT INTO DOCT_F( SLINO                                               ");
                        //                strSql.AppendLine("                            , SEQNO                                              ");
                        //                strSql.AppendLine("                            , FILNM                                              ");
                        //                strSql.AppendLine("                            , CUSER                                              ");
                        //                strSql.AppendLine("                            , CDATE)                                             ");
                        //                strSql.AppendLine("                                                                                 ");
                        //                strSql.AppendLine("                      VALUES('" + _SLINO + "'                                                  ");
                        //                strSql.AppendLine("                            , @SEQNO                                             ");
                        //                strSql.AppendLine("                            , '"+ sFilnm + "'                                                 ");
                        //                strSql.AppendLine("                            , '"+ sUSER + "'                                                 ");
                        //                strSql.AppendLine("                            , CONVERT(VARCHAR(20), GETDATE(), 20))               ");
                        //                strSql.AppendLine("      END");
                        //            }

                        //            cmd.CommandType = CommandType.Text;
                        //            cmd.CommandText = strSql.ToString();
                        //            cmd.ExecuteNonQuery();

                        //            ComnEtcFunc.FTPDirectioryCheck(AttchPath, user, pw);
                        //            ComnEtcFunc.FTPUpload(AttchPath +"/"+ sFilnm, user, pw, FILE);
                        //        }
                        //    }

                        //    DBConn.dbTran.Commit();
                        //    DBConn.dbTran = null;

                        //    XtraMessageBox.Show(resultDt.Rows[0]["MSG"]?.ToString());
                        //    DataRowSendEvent(dicParams);
                        //    DialogResult = DialogResult.OK;
                        //}
                        //catch (Exception ex1)
                        //{
                        //    DBConn.dbTran.Rollback();
                        //    DBConn.dbTran = null;

                        //    XtraMessageBox.Show(ex1.Message);
                        //}
                        #endregion

                        XtraMessageBox.Show(resultDt.Rows[0]["MSG"]?.ToString());

                        if (DataRowSendEvent != null)
                            DataRowSendEvent(dicParams);

                        //DialogResult = DialogResult.OK;
                        Dispose();
                    }

                }

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
              
            }
        }
            #region 엑셀
            Excel.Application ExcelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

        private void EditAndSaveExcelFile(string SavePath)
        {
            uint processId = 0;

            try
            {
                if (!File.Exists(SavePath))
                {
                    XtraMessageBox.Show("엑셀파일 양식이 존재하지 않습니다.");
                    return;
                }

                Cursor = Cursors.WaitCursor;

                ExcelApp = new Excel.Application();
                ExcelApp.DisplayAlerts = false;
                wb = ExcelApp.Workbooks.Open(SavePath //파일경로
                                            , 0 //업데이트링크
                                            , false //읽기전용
                                            , 5 //포멧
                                            , "" //비밀번호
                                            , "" //쓰기비밀번호
                                            , true //읽기전용 권장 메세지 표시x :true
                                            , Excel.XlPlatform.xlWindows //
                                            , "\t" //구분자
                                            , true //편집기능
                                            , false // 알림
                                            , 0 //변환기
                                            , false // 최근사용한 파일목록에 추가
                                            , 1 //
                                            , 0 ); //복구
                int wsCnt = wb.Worksheets.Count;

                string sDOCTP = LkupDOCTP.Text;
                string sDOCOD = LkupDOCTP.EditValue?.ToString();

                DataTable dtCell = new DataTable();

                if (!string.IsNullOrEmpty(sDOCOD))
                {
                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    strSql.AppendLine(" SELECT CELL1      ");//문서번호
                    strSql.AppendLine("      , CELL2      ");//작성자
                    strSql.AppendLine("      , CELL3      ");//부서명
                    strSql.AppendLine("      , CELL4      ");//작성일자
                    strSql.AppendLine("      , CELL5      ");//제목
                    strSql.AppendLine("      , CELL6      ");//거래처명
                    strSql.AppendLine("      , CELL7      ");//금액
                    strSql.AppendLine("   FROM DOCT_K     ");
                    strSql.AppendLine("  WHERE DOCTP = '"+ sDOCOD + "'");

                    dtCell = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                }

                //첫번째 시트에만 적용
                ws = wb.Worksheets[1];
                GetWindowThreadProcessId(new IntPtr(ExcelApp.Hwnd), out processId);

                string sSlino = TxtSLINO.EditValue?.ToString();
                string sTDATE = TDate.DateTime.ToString("yyyy-MM-dd");
                string sUSRNM = LkupUsr.Properties.GetDisplayText(LkupUsr.EditValue);
                string sSTITL = TxtSTITL.EditValue?.ToString();
                string sDEPT = LkupDept.Text?.ToString();
                

                if (dtCell != null)
                {
                    foreach (DataRow row in dtCell.Rows)
                    {
                        string sCELL1 = row["CELL1"]?.ToString();//문서번호
                        string sCELL2 = row["CELL2"]?.ToString();//작성자
                        string sCELL3 = row["CELL3"]?.ToString();//부서명
                        string sCELL4 = row["CELL4"]?.ToString();//작성일자
                        string sCELL5 = row["CELL5"]?.ToString();//제목
                        string sCELL6 = row["CELL6"]?.ToString();//거래처명
                        string sCELL7 = row["CELL7"]?.ToString();//금액
                         
                        if (!string.IsNullOrEmpty(sCELL1)) { ws.Range[sCELL1].Value = sSlino; }
                        if (!string.IsNullOrEmpty(sCELL2)) { ws.Range[sCELL2].Value = sUSRNM; }
                        if (!string.IsNullOrEmpty(sCELL3)) { ws.Range[sCELL3].Value = sDEPT; }
                        if (!string.IsNullOrEmpty(sCELL4)) { ws.Range[sCELL4].Value = sTDATE; }
                        if (!string.IsNullOrEmpty(sCELL5)) { ws.Range[sCELL5].Value = sSTITL; }
                        //if (!string.IsNullOrEmpty(sCELL6)) { _CUSTOM = ws.Range[sCELL6].Value; }
                        //if (!string.IsNullOrEmpty(sCELL7)) { ws.Range[sCELL7].Value = sPay; }
                    }
                }
                
                
                #region 2023-01-06 이전 문서번호,기안자,기안일자, 제목 자동 세팅 코드
                //for (int i = 0; i < wsCnt; i++)
                //{
                //    ws = wb.Worksheets[i + 1];
                //    GetWindowThreadProcessId(new IntPtr(ExcelApp.Hwnd), out processId);

                //    string sSlino = TxtSLINO.EditValue?.ToString();
                //    string sTDATE = TDate.DateTime.ToString("yyyy-MM-dd");
                //    string sUSRNM = LkupUsr.Properties.GetDisplayText(LkupUsr.EditValue);
                //    string sSTITL = TxtSTITL.EditValue?.ToString();
                //    string sDEPT = LkupDept.Text?.ToString();

                //    Excel.Range searchedRange = ExcelApp.get_Range("A1", "BI100");

                //    if (!string.IsNullOrEmpty(sDOCTP))
                //    {
                //        Excel.Range currentFindSlino = null;
                //        Excel.Range currentFindUsrnm = null;
                //        Excel.Range currentFindTdate = null;
                //        Excel.Range currentFindTitle = null;

                //        if (sDOCTP.Equals("기안품의서"))
                //        {
                //            currentFindSlino = searchedRange.Find("문서번호");
                //            currentFindUsrnm = searchedRange.Find("기 안 자");
                //            currentFindTdate = searchedRange.Find("기안일자");
                //            currentFindTitle = searchedRange.Find("제    목");
                //        }
                //        else if (sDOCTP.Equals("지출결의서"))
                //        {
                //            currentFindSlino = searchedRange.Find("문서번호");
                //            currentFindUsrnm = searchedRange.Find("담 당 자");
                //            currentFindTdate = searchedRange.Find("결의일자");
                //            currentFindTitle = searchedRange.Find("제    목");
                //        }

                //        if (currentFindSlino != null)
                //        {
                //            Excel.Range lastMerge = null;
                //            if (currentFindSlino.MergeCells)
                //            {
                //                lastMerge = GetLastMarge(currentFindSlino);
                //            }
                //            else
                //            {
                //                lastMerge = currentFindSlino;
                //            }
                //            ws.Cells[lastMerge.Row, lastMerge.Column + 1] = sSlino;
                //        }
                //        if (currentFindUsrnm != null)
                //        {
                //            Excel.Range lastMerge = null;
                //            if (currentFindSlino.MergeCells)
                //            {
                //                lastMerge = GetLastMarge(currentFindUsrnm);
                //            }
                //            else
                //            {
                //                lastMerge = currentFindUsrnm;
                //            }

                //            ws.Cells[lastMerge.Row, lastMerge.Column + 1] = sUSRNM;
                //        }
                //        if (currentFindTdate != null)
                //        {
                //            Excel.Range lastMerge = null;
                //            if (currentFindSlino.MergeCells)
                //            {
                //                lastMerge = GetLastMarge(currentFindTdate);
                //            }
                //            else
                //            {
                //                lastMerge = currentFindTdate;
                //            }

                //            ws.Cells[lastMerge.Row, lastMerge.Column + 1] = sTDATE;
                //        }
                //        if (currentFindTitle != null)
                //        {
                //            Excel.Range lastMerge = null;
                //            if (currentFindSlino.MergeCells)
                //            {
                //                lastMerge = GetLastMarge(currentFindTitle);
                //            }
                //            else
                //            {
                //                lastMerge = currentFindTitle;
                //            }

                //            ws.Cells[lastMerge.Row, lastMerge.Column + 1] = sSTITL;
                //        }
                //    }
                //}
                #endregion

                wb.SaveAs(SavePath);
                wb.Close(false, Type.Missing, Type.Missing);
                wb = null;
                ExcelApp.Quit();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();

                if (processId != 0)
                {
                    System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)processId);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }
            }
        }
        
        private Excel.Range GetLastMarge(Excel.Range currentFind)
        {
            string sMerge = currentFind.MergeArea.get_Address(Excel.XlReferenceStyle.xlA1);
            string[] sMergeArr = sMerge.Split(':');

            string endMerge = sMergeArr[sMergeArr.Length - 1];

            string col = endMerge.Split('$')[1];
            string row = endMerge.Split('$')[2];

            Excel.Range lastMerge = ws.Range[col + row];

            return lastMerge;
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj); obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally
            {
                GC.Collect();
            }
        }
        #endregion


        private void BtnUpload_Click(object sender, EventArgs e)
        {
            var filePath = string.Empty;

            try
            {
                using (XtraOpenFileDialog openFileDialog = new XtraOpenFileDialog())
                {
                    openFileDialog.Title = "업로드할 파일을 선택해주세요.";
                    openFileDialog.InitialDirectory = "c:\\";
                    //openFileDialog.Filter = "Excel files (*.xls,*xlsx)|*.xls;*xlsx|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 1;
                    openFileDialog.RestoreDirectory = true;
                    openFileDialog.Multiselect = false;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        filePath = openFileDialog.FileName;
                        string filename = Path.GetFileName(filePath); //+ Path.GetExtension(filePath);

                        GridViewRetr.AddNewRow();
                        GridViewRetr.SetFocusedRowCellValue("FILNM", filename);
                        GridViewRetr.SetFocusedRowCellValue("OPATH", filePath);
                        GridViewRetr.UpdateCurrentRow();
                        GridViewRetr.FocusedColumn = GridColFilenm;
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr2_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void DelteBtnFile_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            try
            {
                GridViewRetr.DeleteRow(GridViewRetr.FocusedRowHandle);

                //DataTable dt = GridRetr1.DataSource as DataTable;
                //dt.Rows.RemoveAt(GridViewRetr.FocusedRowHandle);
                //GridRetr1.DataSource = dt;
                //GridViewRetr.UpdateCurrentRow();
                //GridViewRetr.FocusedColumn = GridColFilenm;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GetUploadFileInfo()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT SLINO      ");
            strSql.AppendLine("      , SEQNO      ");
            strSql.AppendLine("      , FILNM      ");
            strSql.AppendLine("      , '' AS OPATH");
            strSql.AppendLine("   FROM DOCT_F     ");
            strSql.AppendLine("  WHERE SLINO = '" + _SLINO + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr1.DataSource = dt;
        }

        private string GetNewDoct()
        {
            string downloadPath = "";
            string sComcd = LkupDOCTP.EditValue?.ToString();

            if (string.IsNullOrEmpty(sComcd))
            {
                XtraMessageBox.Show("문서구분을 선택해주세요.");
                return null;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine("SELECT FILNM                   ");
            strSql.AppendLine("  FROM DOCT_K                  ");
            strSql.AppendLine(" WHERE DOCTP = '" + sComcd + "'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dt == null || dt.Rows.Count <= 0 || string.IsNullOrEmpty(dt.Rows[0]["FILNM"]?.ToString()))
            {
                XtraMessageBox.Show("양식이 없습니다.");
                return null;
            }

            string sFileName = dt.Rows[0]["FILNM"]?.ToString();
            string downPath = tempDoctPath + @"/"+ _tempFileName;

            if (!Directory.Exists(tempDoctPath))
            {
                Directory.CreateDirectory(tempDoctPath);
            }

            if (ComnEtcFunc.FTPFileDownload(sInitDir + sFileName, downPath, user, pw))
            {
                //XtraMessageBox.Show("다운로드를 완료하였습니다.");

                downloadPath = downPath;
                //Process.Start(downPath);
            }

            return downloadPath;
        }

        private string EditDoct()
        {
            string downPath = "";

            byte[] file = null;

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT EXCFIL    ");
            strSql.AppendLine("      , FILNM");
            strSql.AppendLine("   FROM DOCT_M    ");
            strSql.AppendLine("  WHERE SLINO = '" + _SLINO + "'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt != null && dt.Rows.Count > 0 && !DBNull.Value.Equals(dt.Rows[0]["EXCFIL"]))
            {
                downPath = tempDoctPath + @"/"+ _tempFileName;
                file = (byte[])dt.Rows[0]["EXCFIL"];

                if (!Directory.Exists(tempDoctPath))
                {
                    Directory.CreateDirectory(tempDoctPath);
                }
            }

            if (file != null)
            {
                FileStream fs;

                fs = new FileStream(downPath, FileMode.OpenOrCreate, FileAccess.Write);
                fs.Write(file, 0, file.Length);
                fs.Close();

                //XtraMessageBox.Show("다운로드를 완료하였습니다.");
                //Process.Start(downPath);
            }
            else
            {
                downPath = "";
            }

            return downPath;
        }

        private void ExcelCloseAndDelete()
        {
            try
            {
                Process[] prcList = Process.GetProcessesByName("Excel");
                if (prcList.Length != 0)
                {
                    for (int i = 0; i < prcList.Length; i++)
                    {
                        Process prcessInfo = prcList[i];

                        string sProcessFileName = prcessInfo.MainWindowTitle;

                        if (sProcessFileName.Equals(_tempFileName + " - Excel"))
                        {
                            //prcessInfo.Kill();
                            int processId = prcessInfo.Id;

                            if (processId != 0)
                            {
                                System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById(processId);
                                excelProcess.CloseMainWindow();
                                excelProcess.Refresh();
                                excelProcess.Kill();
                            }
                        }
                    }
                }

                Thread.Sleep(1000);

                FileInfo fileInfo = new FileInfo(tempDoctPath + @"/"+ _tempFileName);
                if (fileInfo.Exists)
                {
                    //파일삭제
                    fileInfo.Delete();
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        //문서 신규
        private void BtnNewWrit_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelCloseAndDelete();
                string sDownPath = GetNewDoct();

                if (!string.IsNullOrEmpty(sDownPath))
                {
                    EditAndSaveExcelFile(sDownPath);
                    Process.Start(sDownPath);
                }

            }
            catch (Exception ex)
            {

            }
            
        }

        //문서수정
        private void BtnWritDoct_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelCloseAndDelete();

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "MSTR_RETR");
                dicParams.Add("SLINO", _SLINO);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                string sState = string.Empty;
                string sMSG = string.Empty;

                if (dt.Rows.Count != 0)
                {
                    string sDownPath = tempDoctPath + @"/"+ _tempFileName;
                    //파일다운로드
                    if (!DBNull.Value.Equals(dt.Rows[0]["EXCFIL"]))
                    {
                        byte[] file = (byte[])dt.Rows[0]["EXCFIL"];

                        if (!Directory.Exists(tempDoctPath))
                        {
                            Directory.CreateDirectory(tempDoctPath);
                        }

                        if (file != null)
                        {
                            FileStream fs;

                            fs = new FileStream(sDownPath, FileMode.OpenOrCreate, FileAccess.Write);
                            fs.Write(file, 0, file.Length);
                            fs.Close();
                        }
                    }

                    FileInfo fileInfo = new FileInfo(sDownPath);
                    if (fileInfo.Exists)
                    {
                        EditAndSaveExcelFile(sDownPath);
                        Process.Start(sDownPath);
                    }
                    else
                    {
                        XtraMessageBox.Show("저장된 문서가 없습니다.\r\n지정된 문서 양식을 가져옵니다.");
                        sDownPath = GetNewDoct();

                        if (!string.IsNullOrEmpty(sDownPath))
                        {
                            EditAndSaveExcelFile(sDownPath);
                            Process.Start(sDownPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            
        }

        //문서복사
        private void BtnDoctCopy_Click(object sender, EventArgs e)
        {
            try
            {
                ExcelCloseAndDelete();
                string sDownPath = EditDoct();

                if (!string.IsNullOrEmpty(sDownPath))
                {
                    EditAndSaveExcelFile(sDownPath);
                    Process.Start(sDownPath);
                }

            }
            catch (Exception ex)
            {

            }
        }

        private void RptApplSystemP1_FormClosed(object sender, FormClosedEventArgs e)
        {
            ExcelCloseAndDelete();
        }

        private void LkupDept_EditValueChanged(object sender, EventArgs e)
        {
            string sDept = LkupDept.EditValue?.ToString();

            if (string.IsNullOrEmpty(sDept))
                return;

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT DOCTP AS CD                                  ");
            strSql.AppendLine("      , DOCNM AS NM                                  ");
            strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY LEN(DOCTP), DOCTP) AS SEQ");
            strSql.AppendLine("   FROM DOCT_K                                       ");
            strSql.AppendLine("  WHERE USEDP IN('ALL', '" + sDept + "')");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            dt.Columns["SEQ"].ColumnMapping = MappingType.Hidden; // SEQ 컬럼 안보이게 설정
            ComGrid.SetLookUpEdit(LkupDOCTP, dt, "CD", "NM", "");
        }

        private void LkupDOCTP_EditValueChanged(object sender, EventArgs e)
        {
            string sDoctp = LkupDOCTP.EditValue?.ToString();

            if (string.IsNullOrEmpty(sDoctp))
                return;

            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "GETDOCTNM");
                dicParams.Add("DOCTP", sDoctp);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, this.PROCEDURE_ID, dicParams);

                if(dt != null && dt.Rows.Count > 0)
                {
                    string sFilnm = dt.Rows[0]["FILNM"]?.ToString();

                    _tempFileName = sFilnm;
                }
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void SerchExcelFile(string SavePath, int k)
        {
            uint processId = 0;

            try
            {
                if (!File.Exists(SavePath))
                {
                    XtraMessageBox.Show("엑셀파일 양식이 존재하지 않습니다.");
                    return;
                }

                Cursor = Cursors.WaitCursor;

                ExcelApp = new Excel.Application();
                ExcelApp.DisplayAlerts = false;
                wb = ExcelApp.Workbooks.Open(SavePath //파일경로
                                            , 0 //업데이트링크
                                            , false //읽기전용
                                            , 5 //포멧
                                            , "" //비밀번호
                                            , "" //쓰기비밀번호
                                            , true //읽기전용 권장 메세지 표시x :true
                                            , Excel.XlPlatform.xlWindows //
                                            , "\t" //구분자
                                            , true //편집기능
                                            , false // 알림
                                            , 0 //변환기
                                            , false // 최근사용한 파일목록에 추가
                                            , 1 //
                                            , 0); //복구
                int wsCnt = wb.Worksheets.Count;



                //첫번째 시트에만 적용
                ws = wb.Worksheets[1];
                if (k == 23)  //지출결의서 문서만 실행
                {
                    _CUSTOM = ws.Range["C8"].Value;
                    _PAY = ws.Range["X8"].Value;
                }
                if (k == 37)  //지출결의서 문서만 실행
                {
                    _CUSTOM = ws.Range["C7"].Value;
                    
                }

                wb.SaveAs(SavePath);
                wb.Close(false, Type.Missing, Type.Missing);
                wb = null;
                ExcelApp.Quit();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(ExcelApp);
                GC.Collect();

                if (processId != 0)
                {
                    System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)processId);
                    excelProcess.CloseMainWindow();
                    excelProcess.Refresh();
                    excelProcess.Kill();
                }
            }
        }


    }
}