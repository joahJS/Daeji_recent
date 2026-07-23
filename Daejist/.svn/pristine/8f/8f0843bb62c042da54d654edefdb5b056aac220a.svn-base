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
using MySql.Data.MySqlClient;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;
/*
* 작성일자 : 모름
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
* 
* 수정일자 : 2021-02-25 ~ 2021-02-26
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 그리드 폰트 설정
*/
namespace AccAdm
{
    public partial class AccDeptCdMgtDev : DevExpress.XtraEditors.XtraForm
    {
        public AccDeptCdMgtDev()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void AccDeptCdMgtDev_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            SetLookUpEdit(LkupUpDtCd, "1", "Y", "Y");

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            BtnRetr_Click(null, null);

            //
            
            arrGrdView = new GridView[] { GridViewDeptRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
        }

        private void AccDeptCdMgtDev_FormClosed(object sender, FormClosedEventArgs e)
        {
           
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ClearFps();

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" SELECT A.DEPT_CD ");
            strSql.AppendLine("      , A.DEPT_NM ");
            strSql.AppendLine("      , A.UP_DEPT_CD ");
            strSql.AppendLine("      , (SELECT DEPT_NM ");
            strSql.AppendLine("           FROM ACC_DEPT_CD B ");
            strSql.AppendLine("          WHERE A.UP_DEPT_CD = B.DEPT_CD) AS UP_DEPT_NM ");
            strSql.AppendLine("      , A.USE_YN ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID");
            strSql.AppendLine("   FROM ACC_DEPT_CD A ");
            strSql.AppendLine("  ORDER BY A.DEPT_CD ");


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridDeptRetr.DataSource = dt;

            GridViewDeptRetr_FocusedRowChanged(null, null);
        }

        private void BtnCrte_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            TxtDeptCd.ReadOnly = false;
            TxtDeptCd.Text = null;
            TxtDeptNM.Text = null;
            LkupUpDtCd.EditValue = "";
            RdGbUseYn.EditValue = "Y";
            TxtDeptCd.Enabled = true;
            TxtDeptCd.Focus();
        }

        private bool CheckCodeValueOverlapping()
        {
            bool bYN = false;
            string sDeptCd = TxtDeptCd.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT COUNT(*) AS CNT_VALUE ");
            strSql.AppendLine("   FROM ACC_DEPT_CD ");
            strSql.AppendLine("  WHERE DEPT_CD = '" + sDeptCd + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            double dCntValue = Convert.ToDouble(dt.Rows[0]["CNT_VALUE"]);
            if(dCntValue == 0)
            {
                bYN = false;
            }
            else
            {
                bYN = true;
            }
            return bYN;
        }

        private void TxtDeptCd_Leave(object sender, EventArgs e)
        {
            if (Cursor.Current == BtnCrte.Cursor || Cursor.Current == BtnRetr.Cursor || Cursor.Current == BtnDelete.Cursor
              || Cursor.Current == BtnClose.Cursor || Cursor.Current == BtnSave.Cursor)
            {
                return;
            }
            else
            {
                if (CheckCodeValueOverlapping())
                {
                    XtraMessageBox.Show("해당 코드는 데이터베이스에 이미 존재합니다.\r\n다시 입력하세요.");
                    TxtDeptCd.Focus();
                    TxtDeptCd.SelectAll();
                }
                return;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            string sDeptCd = TxtDeptCd.Text;
            string sDeptNm = TxtDeptNM.Text;
            string sUpDeptCd = LkupUpDtCd.EditValue.ToString();
            string sUseYn = RdGbUseYn.EditValue.ToString();
            string sId = FmMainToolBar2.UserID;

            if (string.IsNullOrEmpty(sDeptCd))
            {
                XtraMessageBox.Show("부서코드를 입력하세요.");
                TxtDeptCd.Focus();
                return;
            }

            if (string.IsNullOrEmpty(sDeptNm))
            {
                XtraMessageBox.Show("부서명을 입력하세요.");
                TxtDeptNM.Focus();
                return;
            }

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            try
            {
                strSql.AppendLine(" IF EXISTS(SELECT* FROM ACC_DEPT_CD WHERE DEPT_CD = '"+ sDeptCd + "')                    ");
                strSql.AppendLine("       BEGIN                                                               ");
                strSql.AppendLine("             UPDATE ACC_DEPT_CD                                            ");
                strSql.AppendLine("                SET DEPT_NM = '"+ sDeptNm + "'                                           ");
                if(!string.IsNullOrEmpty(sUpDeptCd))
                    strSql.AppendLine("              , UP_DEPT_CD = '" + sUpDeptCd + "'                                        ");
                else
                    strSql.AppendLine("              , UP_DEPT_CD = NULL                                        ");
                strSql.AppendLine("                  , USE_YN = '"+ sUseYn + "'                                            ");
                strSql.AppendLine("                  , MFY_DT = CONVERT(VARCHAR(20), GETDATE(), 20)           ");
                strSql.AppendLine("                  , MFY_ID = '"+ sId + "'                                            ");
                strSql.AppendLine("              WHERE DEPT_CD = '"+ sDeptCd + "'                                           ");
                strSql.AppendLine("         END                                                               ");
                strSql.AppendLine("   ELSE                                                                    ");
                strSql.AppendLine("       BEGIN                                                               ");
                strSql.AppendLine("             INSERT INTO ACC_DEPT_CD( DEPT_CD                               ");
                strSql.AppendLine("                                , UP_DEPT_CD                           ");
                strSql.AppendLine("                                    , DEPT_NM                              ");
                strSql.AppendLine("                                    , USE_YN                               ");
                strSql.AppendLine("                                    , ENT_DT                               ");
                strSql.AppendLine("                                    , ENT_ID )                             ");
                strSql.AppendLine("                              VALUES( '"+ sDeptCd + "'                           ");
                if (!string.IsNullOrEmpty(sUpDeptCd))
                    strSql.AppendLine("                                , '"+ sUpDeptCd + "'                           ");
                else
                    strSql.AppendLine("                                , NULL                           ");
                strSql.AppendLine("                                    , '"+ sDeptNm + "'                           ");
                strSql.AppendLine("                                    , '"+ sUseYn + "'                           ");
                strSql.AppendLine("                                    , CONVERT(VARCHAR(20), GETDATE(), 20)  ");
                strSql.AppendLine("                                    , '"+ sId + "')                              ");
                strSql.AppendLine("         END                                                               ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:ACC_DEPT_CD -> DEPT_CD:" + sDeptCd + ",DEPT_NM:" + sDeptNm;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "S", this.Name, sLogRmk, cmd);

                Cursor = Cursors.Default;

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                BtnRetr.PerformClick();
                GridViewDeptRetr.FocusedRowHandle = GridViewDeptRetr.LocateByDisplayText(0, GrdColDeptCd, sDeptCd);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void SetLookUpEdit(DevExpress.XtraEditors.LookUpEdit lkup, string sGb, string sNullYn, string sSetIdx)
        {
            StringBuilder strSql = new StringBuilder();

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS 부서코드 ");
                strSql.AppendLine("      , '' AS 부서명 ");
                strSql.AppendLine("  UNION ALL ");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS 부서코드 ");
                strSql.AppendLine("      , A.DEPT_NM AS 부서명  ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A ");
            }

            strSql.AppendLine("  ORDER BY 부서코드 ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            lkup.Properties.DataSource = dt;
            lkup.Properties.DisplayMember = "부서명";
            lkup.Properties.ValueMember = "부서코드";

            if (sSetIdx.Equals("Y")) lkup.ItemIndex = 0;
        }

        private void ClearFps()
        {
            GridViewDeptRetr.FocusedRowChanged -= GridViewDeptRetr_FocusedRowChanged;
            GridDeptRetr.DataSource = null;
            GridViewDeptRetr.FocusedRowChanged += GridViewDeptRetr_FocusedRowChanged;
        }

        private void GridViewDeptRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            TxtDeptCd.ReadOnly = true;
            TxtDeptCd.Text = GridViewDeptRetr.GetFocusedRowCellValue("DEPT_CD").ToString();
            TxtDeptNM.Text = GridViewDeptRetr.GetFocusedRowCellValue("DEPT_NM").ToString();
            LkupUpDtCd.EditValue = GridViewDeptRetr.GetFocusedRowCellValue("UP_DEPT_CD").ToString();

            RdGbUseYn.EditValue = GridViewDeptRetr.GetFocusedRowCellValue("USE_YN").ToString();

            TxtDeptCd.Enabled = false;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void AccDeptCdMgtDev_KeyDown(object sender, KeyEventArgs e)
       {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnCrte_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnDelete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void GridDeptRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewDeptRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sDeptCd = GridViewDeptRetr.GetFocusedRowCellValue("DEPT_CD")?.ToString();
            string sDeptNm = GridViewDeptRetr.GetFocusedRowCellValue("DEPT_NM")?.ToString();

            if (string.IsNullOrEmpty(sDeptCd))
            {
                XtraMessageBox.Show("부서코드가 존재하지 않습니다.\r\n삭제하려는 데이터를 정확히 클릭하세요.");
                return;
            }

            if (XtraMessageBox.Show("부서코드 : " + sDeptCd + "\r\n부서명 : " + sDeptNm + " \r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
               , "부서코드 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" DELETE FROM ACC_DEPT_CD ");
                strSql.AppendLine("       WHERE DEPT_CD = '" + sDeptCd + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:ACC_DEPT_CD -> DEPT_CD:" + sDeptCd + ",DEPT_NM:" + sDeptNm;
                ClsFunc.LogInsert(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                int idx = GridViewDeptRetr.FocusedRowHandle;
                BtnRetr.PerformClick();
                GridViewDeptRetr.FocusedRowHandle = idx-1;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnExcel_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 엑셀 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string FileName = string.Empty;
            FileDialog fileDlg = new SaveFileDialog();

            try
            {
                string sFileNM = "부서코드 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    GridDeptRetr.ExportToXls(FileName + ".xls");
                    Process.Start(FileName + ".xls");
                }
                fileDlg.Dispose();
            }
            catch (Exception ex)
            {
                if (ex.Message.Equals("Error Opening file"))
                {
                    //파일이 열려있음 체크
                    MessageBox.Show(((ex.InnerException).InnerException).Message);
                }
            }
        }

        private void AccDeptCdMgtDev_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void RdGbUseYn_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                BtnSave.Focus();
            }
        }
    }
}