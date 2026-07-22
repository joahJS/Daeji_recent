using System;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using ComLib;
using System.Data.SqlClient;
using DevExpress.Spreadsheet;
using DevExpress.XtraSplashScreen;

namespace AccAdm
{
    public partial class DoctMgt : DevExpress.XtraEditors.XtraForm
    {
        public DoctMgt()
        {
            InitializeComponent();
        }

        //FTP 서버
        private string sInitDir = ComnEtcFunc.FTP_ROOT + @"/ERP/AprlDocs/";
        private string user = ComnEtcFunc.FTP_USER;
        private string pw = ComnEtcFunc.FTP_PW;

        private void DoctMgt_Load(object sender, EventArgs e)
        {
            this.Icon = ComnEtcFunc.GetFavicon();
            SetLoadFormLayout();

            ComnEtcFunc.gp_SetColorFocused(layoutControl2);

            DataTable dtDoctp = GetDoctpInfo("0");
            dtDoctp.Columns["SEQ"].ColumnMapping = MappingType.Hidden; // SEQ 컬럼 안보이게 설정
            ComGrid.SetLookUpEdit(LkupDoctp, dtDoctp, "CD", "NM", "");
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
        }
        #endregion

        private void DoctMgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            //BtnClose.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
            else if (e.KeyCode == Keys.F1)
                BtnAdd.PerformClick();
            else if (e.KeyCode == Keys.F4)
                BtnDel.PerformClick();
        }

        private void LkupDoctp_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        #region 조회
        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sDOCTP = LkupDoctp.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT A1.DOCTP                                                              ");
            strSql.AppendLine("      , A1.USEDP                                                              ");
            strSql.AppendLine("      , CASE WHEN A1.USEDP = 'ALL' THEN '공용' ELSE B1.DEPT_NM END AS DEPT_NM ");
            strSql.AppendLine("      , A1.DOCNM                                                              ");
            strSql.AppendLine("      , A1.FILNM                                                              ");
            strSql.AppendLine("      , A1.CELL1                                                              ");
            strSql.AppendLine("      , A1.CELL2                                                              ");
            strSql.AppendLine("      , A1.CELL3                                                              ");
            strSql.AppendLine("      , A1.CELL4                                                              ");
            strSql.AppendLine("      , A1.CELL5                                                              ");
            strSql.AppendLine("      , A1.CELL6                                                              ");
            strSql.AppendLine("      , A1.CELL7                                                              ");
            strSql.AppendLine("      , C1.USRNM AS CUSER                                                     ");
            strSql.AppendLine("      , A1.CDATE                                                              ");
            strSql.AppendLine("      , C2.USRNM AS MUSER                                                     ");
            strSql.AppendLine("      , A1.MDATE                                                              ");
            strSql.AppendLine("   FROM DOCT_K A1                                                             ");
            strSql.AppendLine("   LEFT JOIN ACC_DEPT_CD B1                                                   ");
            strSql.AppendLine("     ON A1.USEDP = B1.DEPT_CD                                                 ");
            strSql.AppendLine("   LEFT JOIN ZUSRLST C1                                                       ");
            strSql.AppendLine("     ON A1.CUSER = C1.USRCD                                                   ");
            strSql.AppendLine("   LEFT JOIN ZUSRLST C2                                                       ");
            strSql.AppendLine("     ON A1.MUSER = C2.USRCD                                                   ");

            if (!string.IsNullOrEmpty(sDOCTP))
                strSql.AppendLine("   WHERE DOCTP = '" + sDOCTP + "'");

            strSql.AppendLine(" ORDER BY LEN(A1.DOCTP), A1.DOCTP");


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridRetr.DataSource = dt;

            if (!string.IsNullOrEmpty(_DOCTP))
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColComCd, _DOCTP);

            SetFocuseDoct();

            DataTable dtDoctp = GetDoctpInfo("0");
            dtDoctp.Columns["SEQ"].ColumnMapping = MappingType.Hidden; // SEQ 컬럼 안보이게 설정
            ComGrid.SetLookUpEdit(LkupDoctp, dtDoctp, "CD", "NM", "");
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            SplashScreenManager.ShowForm(typeof(WaitForm1));

            SetFocuseDoct();

            SplashScreenManager.CloseForm();
        }

        private void SetFocuseDoct()
        {
            try
            {
                string focusedFileName = GridViewRetr.GetFocusedRowCellValue("FILNM")?.ToString();

                if (string.IsNullOrEmpty(focusedFileName))
                {
                    int sheetCnt = spread.ActiveWorksheet.Workbook.Worksheets.Count;

                    spread.ActiveWorksheet.Workbook.Worksheets.Add("EmptySheet1");
                    for (int i = 0; i < sheetCnt; i++)
                    {
                        spread.ActiveWorksheet.Workbook.Worksheets.RemoveAt(0);
                    }

                    return;
                }

                byte[] DownloadFile = ComnEtcFunc.DownloadFTPFile_ByteArray(sInitDir + focusedFileName, user, pw);

                if (DownloadFile == null)
                {
                    int sheetCnt = spread.ActiveWorksheet.Workbook.Worksheets.Count;

                    spread.ActiveWorksheet.Workbook.Worksheets.Add("EmptySheet1");
                    for (int i = 0; i < sheetCnt; i++)
                    {
                        spread.ActiveWorksheet.Workbook.Worksheets.RemoveAt(0);
                    }

                    return;
                }

                spread.LoadDocument(DownloadFile, DocumentFormat.Xlsm);
            }
            catch (Exception ex)
            {

                if (ex.Message.Equals("Cannot rename a sheet to the same name as existing sheet."))
                {

                }
                else
                {
                    XtraMessageBox.Show(ex.Message);
                }
            }
        }
        #endregion

        #region 추가
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            PopDocFile frm = new PopDocFile();
            frm.Owner = this;
            frm.AddModifyGb = "ADD";
            frm.DataRowSendEvent += new PopDocFile.SendDataHandler(GetAddData);

            if (frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr.PerformClick();
            }
        }

        private string _DOCTP = string.Empty;
        private void GetAddData(string sDoctp)
        {
            _DOCTP = sDoctp;
        }
        #endregion

        #region 삭제
        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            int[] selRows = GridViewRetr.GetSelectedRows();

            if (selRows.Length == 0)
            {
                XtraMessageBox.Show("삭제할 항목을 선택해주세요.");
                return;
            }

            if (XtraMessageBox.Show("선택한" + selRows.Length + "개의 항목을 삭제하시겠습니까?", "삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                StringBuilder strSql = new StringBuilder();

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                DataTable dt = (DataTable)GridRetr.DataSource;

                for (int i = 0; i < selRows.Length; i++)
                {
                    //정렬 후 변경된 index값 원래 index값으로 변경
                    int orignalIndex = GridViewRetr.GetDataSourceRowIndex(selRows[i]);

                    DataRow row = dt.Rows[orignalIndex];

                    string sDOCTP = row["DOCTP"]?.ToString();

                    //파일삭제
                    string sBefoNm = row["FILNM"]?.ToString();
                    if (!string.IsNullOrEmpty(sBefoNm))
                    {
                        ComnEtcFunc.DeleteFTPFile(sInitDir + sBefoNm, user, pw);
                    }

                    strSql.Clear();
                    strSql.AppendLine(" DELETE FROM DOCT_K");
                    strSql.AppendLine("  WHERE DOCTP = '" + sDOCTP + "'  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = selRows[0] - 1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 수정
        private void GridViewRetr_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 수정 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                PopDocFile frm = new PopDocFile();
                frm.Owner = this;
                frm._DOCTP = GridViewRetr.GetFocusedRowCellValue("DOCTP")?.ToString();
                frm.AddModifyGb = "MOD";
                frm.PDataRow = GridViewRetr.GetFocusedDataRow();
                frm.DataRowSendEvent += new PopDocFile.SendDataHandler(GetAddData);

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    BtnRetr.PerformClick();
                }
            }
        }
        #endregion

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }



        #region LOOKUP DATA
        private DataTable GetDoctpInfo(string sGB)
        {
            StringBuilder strSql = new StringBuilder();

            if (sGB.Equals("0"))
            {
                strSql.Clear();
                strSql.AppendLine(" SELECT '' AS CD    ");
                strSql.AppendLine("      , '' AS NM    ");
                strSql.AppendLine("      , 0 AS SEQ    ");
                strSql.AppendLine("  UNION ALL         ");
                strSql.AppendLine(" SELECT DOCTP AS CD                                  ");
                strSql.AppendLine("      , DOCNM AS NM                                  ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY LEN(DOCTP), DOCTP)");
                strSql.AppendLine("   FROM DOCT_K                                       ");
            }

            return DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
        }
        #endregion

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void DoctMgt_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];
                string path = ComnEtcFunc.GetLayoutPath();
                if (!System.IO.Directory.Exists(path))
                {
                    System.IO.Directory.CreateDirectory(path);
                }

                layoutControl1.SaveLayoutToXml(path + @"\" + this.Name + "_Layout.xaml");

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }
    }
}