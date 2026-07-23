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
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;
using System.Globalization;
/*
* 작성일자 : 모름
* 작성자 : 고혜성
* ---------------------HISTORY-----------------------
* 
* 수정일자 : 2021-02-25 ~ 2021-02-26
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 그리드 폰트 설정
*            2. 레이아웃 전체 저장 설정
*/
namespace AccAdm
{
    public partial class HrCertificateOfEmployment : DevExpress.XtraEditors.XtraForm
    {
        public HrCertificateOfEmployment()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        private GridView[] arrGrdView;
        private void HrCertificateOfEmployment_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "HrCertificateOfEmployment");

            Cursor = Cursors.WaitCursor;

            DataTable dtDeptCD = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupDeptCd, dtDeptCD, GridIssue, BGridColDeptCd, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupDeptCd, dtDeptCD, GridHis, BGridColHisDeptCd, "CD", "NM", "");
            ComLib.ComGrid.SetLookUpEdit(LkupDeptCd, dtDeptCD, "CD", "NM", "Y");

            DataTable dtGradeCD = GetLookUpData("2", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupGradeCD, dtGradeCD, GridIssue, BGridColGrade, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupGradeCD, dtGradeCD, GridHis, BGridColHisGradeCd, "CD", "NM", "");
            ComLib.ComGrid.SetLookUpEdit(LkupGradeCd, dtGradeCD, "CD", "NM", "Y");

            Cursor = Cursors.Default;

            ComLib.ClsFunc.SetGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "HrCertificateOfEmployment", BGridViewIssue);
            ComLib.ClsFunc.SetGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "HrCertificateOfEmployment1", BGridViewHis);

            DataTable dt = GetLookUpData("3","A","A");
            ComGrid.SetLookUpEdit(LkupEmpNm, dt, "CD", "CD", "Y");
            LkupEmpNm.Properties.PopulateColumns();
            LkupEmpNm.Properties.Columns[1].Visible = false;

            arrGrdView = new GridView[] { BGridViewIssue, BGridViewHis, GridViewEmp };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            GetEmployeeBasisInfo();
        }
        
        #region[LookUpData Setting]
        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("A"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '공용' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , A.DEPT_CD AS SEQ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  WHERE USE_YN = 'Y' ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'GRADE_CD'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT EMP_NM AS CD");
                strSql.AppendLine(" 	  , '' AS NM   ");
                strSql.AppendLine(" FROM HR_EMP_BASIS  ");
                strSql.AppendLine(" WHERE EMPL_GB = '"+ RdgbEmplGB.EditValue.ToString() + "'");
                strSql.AppendLine(" GROUP BY EMP_NM    ");
            }

            if (sParam.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }
            else if (sParam.Equals("A"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY CD");
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            return dt;
        }

        #endregion[LookUpData Setting]
        
        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GetEmployeeBasisInfo();
        }

        private void GetEmployeeBasisInfo()
        {
            Cursor = Cursors.WaitCursor;

            string sEmpNm = LkupEmpNm.EditValue?.ToString();
            string sEmplGB = RdgbEmplGB.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , A.EMP_NM ");
            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
            strSql.AppendLine("  WHERE 1 = 1 ");
            if (!string.IsNullOrEmpty(sEmplGB)) strSql.AppendLine("    AND EMPL_GB = '" + sEmplGB + "' ");
            if (!string.IsNullOrEmpty(sEmpNm)) strSql.AppendLine("    AND EMP_NM = '" + sEmpNm + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridEmp.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;

            string sYearInNow = DateTime.Today.Year.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT ISNULL(MAX(ISSUE_SEQ), 0) + 1 AS MAX_VALUE ");
            strSql.AppendLine("   FROM HR_CERT_EMPT  ");
            strSql.AppendLine("  WHERE ISSUE_YEAR = '" + sYearInNow + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            string sMaxValue = dt.Rows[0]["MAX_VALUE"]?.ToString();

            TxtIssueYear.EditValue = sYearInNow;
            TxtIssueSeq.EditValue = sMaxValue;

            string sEmpId = GridViewEmp.GetFocusedRowCellValue("EMP_ID")?.ToString();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_NM ");
            strSql.AppendLine(" 	 , A.CHN_NM ");
            strSql.AppendLine(" 	 , A.ENTRANCE_YMD ");
            strSql.AppendLine(" 	 , A.DEPT_CD ");
            strSql.AppendLine(" 	 , A.GRADE_CD ");
            strSql.AppendLine(" 	 , A.IDT_NO ");
            strSql.AppendLine(" 	 , B.ADDR ");
            strSql.AppendLine(" 	 , B.DTL_ADDR ");
            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
            strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_PERSONAL B  ");
            strSql.AppendLine("     ON B.EMP_ID = A.EMP_ID ");
            strSql.AppendLine("  WHERE A.EMP_ID = '" + sEmpId + "' ");

            DataTable dtInfo = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dtInfo.Rows.Count > 0)
            {
                string sEmpNm = dtInfo.Rows[0]["EMP_NM"]?.ToString();
                string sChnNm = dtInfo.Rows[0]["CHN_NM"]?.ToString();
                string sEntranceYmd = dtInfo.Rows[0]["ENTRANCE_YMD"]?.ToString();
                string sDeptCd = dtInfo.Rows[0]["DEPT_CD"]?.ToString();
                string sGradeCd = dtInfo.Rows[0]["GRADE_CD"]?.ToString();
                string sIdtNo = dtInfo.Rows[0]["IDT_NO"]?.ToString();
                if (!string.IsNullOrEmpty(sIdtNo))
                {
                    string sIdtNo2 = ComnEtcFunc.Decrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);

                    if (!string.IsNullOrEmpty(sIdtNo2))
                    {
                        sIdtNo = sIdtNo2;
                    }
                }

                string sAddr = dtInfo.Rows[0]["ADDR"]?.ToString();
                string sDtlAddr = dtInfo.Rows[0]["DTL_ADDR"]?.ToString();

                TxtEmpKor.EditValue = sEmpNm;
                TxtEmpChn.EditValue = sChnNm;
                
                if(sEntranceYmd.Length == 8)
                {
                    DateEditEntranceYmd.EditValue = sEntranceYmd.Substring(0, 4) + "-" + sEntranceYmd.Substring(4, 2) + "-" + sEntranceYmd.Substring(6, 2);
                }

                DateEditIssueYmd.EditValue = DateTime.Today;
                DateEditFromTo.EditValue = DateTime.Today;

                LkupDeptCd.EditValue = sDeptCd;
                LkupGradeCd.EditValue = sGradeCd;

                TxtIdtNo.EditValue = sIdtNo;
                TxtIdtNo2.EditValue = sIdtNo;
                TxtAddr.EditValue = sAddr + " " + sDtlAddr;
                TxtPurp.EditValue = string.Empty;

                TxtAddMfyGb.EditValue = "ADD";
            }
            Cursor = Cursors.Default;

            TxtPurp.Focus();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sIssueYear = TxtIssueYear.EditValue?.ToString();
            string sIssueSeq = TxtIssueSeq.EditValue?.ToString();
            string sEmpId = TxtEmpId.EditValue?.ToString();
            string sEmpNm = TxtEmpKor.EditValue?.ToString();
            string sChnNm = TxtEmpChn.EditValue?.ToString();
            string sIssueYmd = DateEditIssueYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sEntranceYmd = DateEditEntranceYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sNowYmd = DateEditFromTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sDeptCd = LkupDeptCd.EditValue?.ToString();
            string sGradeCd = LkupGradeCd.EditValue?.ToString();
            string sIdtNo = TxtIdtNo2.EditValue?.ToString();
            if (!string.IsNullOrEmpty(sIdtNo))
            {
                sIdtNo = ComnEtcFunc.Encrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);
            }

            string sAddr = TxtAddr.EditValue?.ToString();
            string sIssuePPs = TxtPurp.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;
                
                string sAddMfyGb = TxtAddMfyGb.EditValue?.ToString();
                if (sAddMfyGb.Equals("MODIFY"))
                {
                    DataRow drInfo = BGridViewIssue.GetFocusedDataRow();
                    SaveHistory(cmd, drInfo);
                }

                strSql.Clear();
                strSql.AppendLine(" ");

                #region mariaDB
                //strSql.AppendLine(" INSERT INTO HR_CERT_EMPT ");
                //strSql.AppendLine("           ( ");
                //strSql.AppendLine("             ISSUE_YEAR    , ISSUE_SEQ, EMP_ID ");
                //strSql.AppendLine("           , EMP_NM        , CHN_NM   , ISSUE_YMD ");
                //strSql.AppendLine("           , ENTRANCE_YMD  , NOW_YMD  , DEPT_CD ");
                //strSql.AppendLine("           , JOBPOSITION_CD, GRADE_CD , IDT_NO ");
                //strSql.AppendLine("           , ADDR          , ISSUE_PPS, ENT_DT ");
                //strSql.AppendLine("           , ENT_ID        , MFY_DT   , MFY_ID ");
                //strSql.AppendLine("           ) ");
                //strSql.AppendLine("      VALUES ");
                //strSql.AppendLine("           ( ");
                //strSql.AppendLine("             '" + sIssueYear   + "',  " + sIssueSeq + " ,  " + sEmpId    + " ");
                //strSql.AppendLine("           , '" + sEmpNm       + "', '" + sChnNm    + "', '" + sIssueYmd + "' ");
                //strSql.AppendLine("           , '" + sEntranceYmd + "', '" + sNowYmd   + "', '" + sDeptCd   + "' ");
                //strSql.AppendLine("           , '" + sGradeCd     + "', '" + sGradeCd  + "', '" + sIdtNo    + "' ");
                //strSql.AppendLine("           , '" + sAddr        + "', '" + sIssuePPs + "', NOW() ");
                //strSql.AppendLine("           , ''                    ,               NOW(), '' ");
                //strSql.AppendLine("           ) ");
                //strSql.AppendLine(" 		 ON DUPLICATE KEY UPDATE ");
                //strSql.AppendLine("             EMP_NM         = '" + sEmpNm       + "', CHN_NM    = '" + sChnNm    + "', ISSUE_YMD = '" + sIssueYmd + "' ");
                //strSql.AppendLine("           , ENTRANCE_YMD   = '" + sEntranceYmd + "', NOW_YMD   = '" + sNowYmd   + "', DEPT_CD   = '" + sDeptCd   + "' ");
                //strSql.AppendLine("           , JOBPOSITION_CD = '" + sGradeCd     + "', GRADE_CD  = '" + sGradeCd  + "', IDT_NO    = '" + sIdtNo    + "' ");
                //strSql.AppendLine("           , ADDR           = '" + sAddr        + "', ISSUE_PPS = '" + sIssuePPs + "', MFY_DT    = NOW() ");
                //strSql.AppendLine("           , MFY_ID         = '' ");
                #endregion

                strSql.AppendLine(" IF EXISTS(SELECT* FROM HR_CERT_EMPT WHERE ISSUE_YEAR = '"+ sIssueYear + "' AND ISSUE_SEQ = "+ sIssueSeq + " AND EMP_ID = "+ sEmpId + ")");
                strSql.AppendLine("    BEGIN                                                                                   ");
                strSql.AppendLine("          UPDATE HR_CERT_EMPT                                                               ");
                strSql.AppendLine("             SET EMP_NM = '" + sEmpNm       + "', CHN_NM = '" + sChnNm    + "', ISSUE_YMD = '" + sIssueYmd + "' ");
	            strSql.AppendLine("               , ENTRANCE_YMD = '" + sEntranceYmd + "', NOW_YMD = '" + sNowYmd   + "', DEPT_CD = '" + sDeptCd   + "' ");
	            strSql.AppendLine("               , JOBPOSITION_CD = '" + sGradeCd     + "', GRADE_CD = '" + sGradeCd  + "', IDT_NO = '" + sIdtNo    + "' ");
	            strSql.AppendLine("               , ADDR = '" + sAddr        + "', ISSUE_PPS = '" + sIssuePPs + "', MFY_DT = GETDATE() ");
	            strSql.AppendLine("               , MFY_ID = '' ");
                strSql.AppendLine("           WHERE ISSUE_YEAR = '"+ sIssueYear + "'");
                strSql.AppendLine("             AND ISSUE_SEQ = "+ sIssueSeq + "  ");
                strSql.AppendLine("             AND EMP_ID = "+ sEmpId + "     ");
                strSql.AppendLine("      END                       ");
                strSql.AppendLine(" ELSE                           ");
                strSql.AppendLine("    BEGIN                       ");
                strSql.AppendLine("          INSERT INTO HR_CERT_EMPT ");
                strSql.AppendLine("               (");
                strSql.AppendLine("                 ISSUE_YEAR, ISSUE_SEQ, EMP_ID ");
                strSql.AppendLine("               , EMP_NM, CHN_NM, ISSUE_YMD ");
                strSql.AppendLine("               , ENTRANCE_YMD, NOW_YMD, DEPT_CD ");
                strSql.AppendLine("               , JOBPOSITION_CD, GRADE_CD, IDT_NO ");
                strSql.AppendLine("               , ADDR, ISSUE_PPS, ENT_DT ");
                strSql.AppendLine("               , ENT_ID ");
                strSql.AppendLine("               ) ");
                strSql.AppendLine("          VALUES ");
                strSql.AppendLine("               (");
                strSql.AppendLine("                 '" + sIssueYear   + "', " + sIssueSeq + ", " + sEmpId    + " ");
                strSql.AppendLine("               , '" + sEmpNm       + "', '" + sChnNm    + "', '" + sIssueYmd + "' ");
                strSql.AppendLine("               , '" + sEntranceYmd + "', '" + sNowYmd   + "', '" + sDeptCd   + "' ");
                strSql.AppendLine("               , '" + sGradeCd     + "', '" + sGradeCd  + "', '" + sIdtNo    + "' ");
                strSql.AppendLine("               , '" + sAddr        + "', '" + sIssuePPs + "', GETDATE() ");
                strSql.AppendLine("               , ''");
                strSql.AppendLine("               ) ");
                strSql.AppendLine("      END");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                Cursor = Cursors.Default;

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                BtnRetr_Click(null, null);
            }
            catch(Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void SaveHistory(SqlCommand cmd, DataRow drInfo)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();
            
            string sIssueYear = drInfo["ISSUE_YEAR"]?.ToString();
            string sIssueSeq = drInfo["ISSUE_SEQ"]?.ToString();
            string sOccurDt = DateTime.Today.ToString().Replace("-", "").Substring(0, 8);
            string sEmpId = drInfo["EMP_ID"]?.ToString();
            string sEmpNm = drInfo["EMP_NM"]?.ToString();
            string sChnNm = drInfo["CHN_NM"]?.ToString();
            string sIssueYmd = drInfo["ISSUE_YMD"]?.ToString();
            string sEntraceYmd = drInfo["ENTRANCE_YMD"]?.ToString();
            string sNowYmd = drInfo["NOW_YMD"]?.ToString();
            string sDeptCd = drInfo["DEPT_CD"]?.ToString();
            string sJobPosition = drInfo["JOBPOSITION_CD"]?.ToString();
            string sGradeCd = drInfo["GRADE_CD"]?.ToString();
            string sIdtNo = drInfo["IDT_NO"]?.ToString();
            if (!string.IsNullOrEmpty(sIdtNo))
            {
                sIdtNo = ComnEtcFunc.Encrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);
            }
            string sAddr = drInfo["ADDR"]?.ToString();
            string sIssuePps = drInfo["ISSUE_PPS"]?.ToString();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT ISNULL(MAX(SEQ), 0) + 1 AS MAX_VALUE ");
            strSql.AppendLine("   FROM HR_CERT_EMPT_HISTORY ");
            strSql.AppendLine("  WHERE ISSUE_YEAR = '" + sIssueYear + "' ");
            strSql.AppendLine("    AND ISSUE_SEQ = " + sIssueSeq + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt.Rows.Count > 0)
            {
                string sMaxValue = dt.Rows[0]["MAX_VALUE"]?.ToString();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" INSERT INTO HR_CERT_EMPT_HISTORY ");
                strSql.AppendLine("           ( ");
                strSql.AppendLine("             ISSUE_YEAR, ISSUE_SEQ, SEQ ");
                strSql.AppendLine("           , OCCUR_DT, MFY_ID, EMP_ID ");
                strSql.AppendLine("           , EMP_NM, CHN_NM, ISSUE_YMD ");
                strSql.AppendLine("           , ENTRANCE_YMD, NOW_YMD, DEPT_CD ");
                strSql.AppendLine("           , JOBPOSITION_CD, GRADE_CD, IDT_NO ");
                strSql.AppendLine("           , ADDR, ISSUE_PPS ");
                strSql.AppendLine("           ) ");
                strSql.AppendLine("      VALUES ");
                strSql.AppendLine("           (  ");
                strSql.AppendLine("             '" + sIssueYear   + "', " + sIssueSeq  + ", " + sMaxValue   + " ");
                strSql.AppendLine("           , '" + sOccurDt     + "',                 '', " + sEmpId      + " ");
                strSql.AppendLine("           , '" + sEmpNm       + "', '" + sChnNm    + "', '" + sIssueYmd + "' ");
                strSql.AppendLine("           , '" + sEntraceYmd  + "', '" + sNowYmd   + "', '" + sDeptCd   + "' ");
                strSql.AppendLine("           , '" + sJobPosition + "', '" + sGradeCd  + "', '" + sIdtNo    + "' ");
                strSql.AppendLine("           , '" + sAddr        + "', '" + sIssuePps + "' ");
                strSql.AppendLine("           ) ");

                Cursor = Cursors.Default;

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();
            }
            
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("EXCEL", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 프린트 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (BGridViewIssue.RowCount > 0)
            {
                string sIssueYear = BGridViewIssue.GetFocusedRowCellValue("ISSUE_YEAR")?.ToString();
                string sIssueSeq = BGridViewIssue.GetFocusedRowCellValue("ISSUE_SEQ")?.ToString();
                string sIssueCont = "제 " + sIssueYear + "-" + sIssueSeq + "호";

                string sEmpNm = BGridViewIssue.GetFocusedRowCellValue("EMP_NM")?.ToString();
                string sDeptNm = BGridViewIssue.GetFocusedRowCellDisplayText("DEPT_CD")?.ToString();
                string sIdtNo = BGridViewIssue.GetFocusedRowCellValue("IDT_NO")?.ToString();
                string sAddr = BGridViewIssue.GetFocusedRowCellValue("ADDR")?.ToString();

                string sEntranceYmd = BGridViewIssue.GetFocusedRowCellValue("ENTRANCE_YMD")?.ToString();
                string sEntranceResult = string.Empty;
                if (sEntranceYmd.Length == 8)
                {
                    sEntranceResult = sEntranceYmd.Substring(0, 4) + "-" + sEntranceYmd.Substring(4, 2) + "-" + sEntranceYmd.Substring(6, 2);
                }

                string sGradeNm = BGridViewIssue.GetFocusedRowCellDisplayText("GRADE_CD")?.ToString();
                string sPurp = BGridViewIssue.GetFocusedRowCellValue("ISSUE_PPS")?.ToString();

                string sIssueYmd = BGridViewIssue.GetFocusedRowCellValue("ISSUE_YMD")?.ToString();
                string sIssueResult = string.Empty;
                if (sIssueYmd.Length == 8)
                {
                    sIssueResult = sIssueYmd.Substring(0, 4) + "년" + sIssueYmd.Substring(4, 2) + "월" + sIssueYmd.Substring(6, 2) + "일";
                }

                DataTable dt = new DataTable();

                dt.TableName = "INFO";
                dt.Columns.Add("ISSUE_CONTENT");
                dt.Columns.Add("EMP_NM");
                dt.Columns.Add("DEPT_NM");
                dt.Columns.Add("IDT_NO");
                dt.Columns.Add("ADDR");
                dt.Columns.Add("ENTRANCE_YMD");
                dt.Columns.Add("GRADE_NM");
                dt.Columns.Add("PURP");
                dt.Columns.Add("ISSUE_YMD");

                DataRow row = dt.NewRow();

                row["ISSUE_CONTENT"] = sIssueCont;
                row["EMP_NM"] = sEmpNm;
                row["DEPT_NM"] = sDeptNm;
                row["IDT_NO"] = sIdtNo;
                row["ADDR"] = sAddr;
                row["ENTRANCE_YMD"] = sEntranceResult;
                row["GRADE_NM"] = sGradeNm;
                row["PURP"] = sPurp;
                row["ISSUE_YMD"] = sIssueResult;

                dt.Rows.Add(row);

                ReportViewer fm = new ReportViewer(dt, "RptCertOfEmp");
                fm.ShowDialog();
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridViewEmp_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            ClearAllComponents();

            /*
             * 2021-01-21 현업요청
             * 주민등록번호 뒷자리는 옵션으로 보이거나 숨길수있도록 수정
             */

            string sAppearYn = RdgbAppearYn.EditValue?.ToString();
            
            string sEmpId = GridViewEmp.GetFocusedRowCellValue("EMP_ID")?.ToString();
            TxtEmpId.EditValue = sEmpId;
            if (!string.IsNullOrEmpty(sEmpId))
            {
                Cursor = Cursors.WaitCursor;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.ISSUE_YEAR ");
                strSql.AppendLine("      , A.ISSUE_SEQ ");
                strSql.AppendLine("      , A.EMP_ID ");
                strSql.AppendLine("      , A.EMP_NM ");
                strSql.AppendLine("      , A.CHN_NM ");
                strSql.AppendLine("      , A.ISSUE_YMD ");
                strSql.AppendLine("      , A.ENTRANCE_YMD ");
                strSql.AppendLine("      , A.NOW_YMD ");
                strSql.AppendLine("      , A.DEPT_CD ");
                strSql.AppendLine("      , A.JOBPOSITION_CD ");
                strSql.AppendLine("      , A.GRADE_CD ");
                strSql.AppendLine("      , A.IDT_NO");
                strSql.AppendLine("      , A.IDT_NO AS IDT_NO2 ");
                //strSql.AppendLine("      , CASE WHEN A.IDT_NO IS NULL OR A.IDT_NO = '' THEN ''                                                     ");
                //strSql.AppendLine("             WHEN '" + sAppearYn + "' = 'Y' THEN A.IDT_NO                                                       ");
                //strSql.AppendLine("             ELSE SUBSTRING(A.IDT_NO, 1, 8) + REPLICATE('*', 14 - LEN(SUBSTRING(A.IDT_NO, 1, 8))) END AS IDT_NO ");
                strSql.AppendLine("      , A.ADDR ");
                strSql.AppendLine("      , A.ISSUE_PPS ");
                strSql.AppendLine("   FROM HR_CERT_EMPT A ");
                strSql.AppendLine("  WHERE EMP_ID = '" + sEmpId + "' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                
                if(dt != null && dt.Rows.Count > 0)
                {
                    foreach(DataRow row in dt.Rows)
                    {
                        string sIdtNo = row["IDT_NO"]?.ToString();
                        string sIssuYMd = row["ISSUE_YMD"]?.ToString();

                        if (!string.IsNullOrEmpty(sIssuYMd))
                        {
                            DateTime dtIssuYMD = DateTime.ParseExact(sIssuYMd, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None);
                            DateTime editDate = new DateTime(2023, 02, 07);

                            if (!string.IsNullOrEmpty(sIdtNo) && dtIssuYMD >= editDate)
                            {
                                string sIdtNo2 = ComnEtcFunc.Decrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);

                                if (!string.IsNullOrEmpty(sIdtNo2))
                                {
                                    sIdtNo = sIdtNo2;
                                }

                                row["IDT_NO2"] = sIdtNo;

                                if (sAppearYn.Equals("N"))
                                {
                                    row["IDT_NO"] = sIdtNo.Substring(0, 8) + "******";
                                }
                                else
                                {
                                    row["IDT_NO"] = sIdtNo;
                                }
                            }
                        }
                    }

                    string sYear = dt.Rows[0]["ISSUE_YEAR"]?.ToString();
                    string sSeq = dt.Rows[0]["ISSUE_SEQ"]?.ToString();
                    GetCertificationHistoryInfo(sYear, sSeq);

                    setEditFocuse();
                }
                else
                {
                    GridHis.DataSource = null;
                }

                GridIssue.DataSource = dt;

                Cursor = Cursors.Default;
            }
        }

        private void ClearAllComponents()
        {
            TxtEmpId.EditValue = string.Empty;
            TxtIssueYear.EditValue = string.Empty;
            TxtIssueSeq.EditValue = string.Empty;
            TxtEmpKor.EditValue = string.Empty;
            TxtEmpChn.EditValue = string.Empty;
            TxtIdtNo.EditValue = string.Empty;
            TxtAddr.EditValue = string.Empty;
            LkupDeptCd.EditValue = string.Empty;
            LkupGradeCd.EditValue = string.Empty;
            DateEditEntranceYmd.EditValue = string.Empty;
            DateEditFromTo.EditValue = string.Empty;
            DateEditIssueYmd.EditValue = string.Empty;
            TxtAddMfyGb.EditValue = string.Empty;
        }

        private void BGridViewIssue_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if(BGridViewIssue.RowCount > 0)
            {
                setEditFocuse();
            }
            else
            {
                ClearComponents();
            }
        }

        private void setEditFocuse()
        {
            string sIssueYear = BGridViewIssue.GetFocusedRowCellValue("ISSUE_YEAR")?.ToString();
            string sIssueSeq = BGridViewIssue.GetFocusedRowCellValue("ISSUE_SEQ")?.ToString();
            string sEmpNm = BGridViewIssue.GetFocusedRowCellValue("EMP_NM")?.ToString();
            string sChnNm = BGridViewIssue.GetFocusedRowCellValue("CHN_NM")?.ToString();

            string sIssueYmd = BGridViewIssue.GetFocusedRowCellValue("ISSUE_YMD")?.ToString();
            string sIssueYmdResult = string.Empty;
            if (sIssueYmd != null && sIssueYmd.Length == 8)
            {
                sIssueYmdResult = sIssueYmd.Substring(0, 4) + "-" + sIssueYmd.Substring(4, 2) + "-" + sIssueYmd.Substring(6, 2);
            }

            string sEntranceYmd = BGridViewIssue.GetFocusedRowCellValue("ENTRANCE_YMD")?.ToString();
            string sEntranceYmdResult = string.Empty;
            if (sEntranceYmd != null && sEntranceYmd.Length == 8)
            {
                sEntranceYmdResult = sEntranceYmd.Substring(0, 4) + "-" + sEntranceYmd.Substring(4, 2) + "-" + sEntranceYmd.Substring(6, 2);
            }

            string sNowYmd = BGridViewIssue.GetFocusedRowCellValue("NOW_YMD")?.ToString();
            string sNowYmdResult = string.Empty;
            if (sNowYmd!= null && sNowYmd.Length == 8)
            {
                sNowYmdResult = sNowYmd.Substring(0, 4) + "-" + sNowYmd.Substring(4, 2) + "-" + sNowYmd.Substring(6, 2);
            }

            string sDeptCd = BGridViewIssue.GetFocusedRowCellValue("DEPT_CD")?.ToString();
            string sGradeCd = BGridViewIssue.GetFocusedRowCellValue("GRADE_CD")?.ToString();
            string sIdtNo = BGridViewIssue.GetFocusedRowCellValue("IDT_NO")?.ToString();
            string sAddr = BGridViewIssue.GetFocusedRowCellValue("ADDR")?.ToString();
            string sIssuePPS = BGridViewIssue.GetFocusedRowCellValue("ISSUE_PPS")?.ToString();

            TxtIssueYear.EditValue = sIssueYear;
            TxtIssueSeq.EditValue = sIssueSeq;
            TxtEmpKor.EditValue = sEmpNm;
            TxtEmpChn.EditValue = sChnNm;
            DateEditIssueYmd.EditValue = sIssueYmdResult;
            DateEditEntranceYmd.EditValue = sEntranceYmdResult;
            DateEditFromTo.EditValue = sNowYmdResult;
            LkupDeptCd.EditValue = sDeptCd;
            LkupGradeCd.EditValue = sGradeCd;
            TxtIdtNo.EditValue = sIdtNo;
            TxtAddr.EditValue = sAddr;
            TxtPurp.EditValue = sIssuePPS;

            if(!string.IsNullOrEmpty(sIssueYear) && !string.IsNullOrEmpty(sIssueSeq))
            {
                GetCertificationHistoryInfo(sIssueYear, sIssueSeq);

                TxtAddMfyGb.EditValue = "MODIFY";
            }
        }

        private void ClearComponents()
        {
            Cursor = Cursors.WaitCursor;

            TxtIssueYear.EditValue = string.Empty;
            TxtIssueSeq.EditValue = string.Empty;
            TxtEmpKor.EditValue = string.Empty;
            TxtEmpChn.EditValue = string.Empty;
            DateEditIssueYmd.EditValue = string.Empty;
            DateEditEntranceYmd.EditValue = string.Empty;
            DateEditFromTo.EditValue = string.Empty;
            LkupDeptCd.EditValue = string.Empty;
            LkupGradeCd.EditValue = string.Empty;
            TxtIdtNo.EditValue = string.Empty;
            TxtAddr.EditValue = string.Empty;
            TxtPurp.EditValue = string.Empty;

            Cursor = Cursors.Default;
        }

        private void GetCertificationHistoryInfo(string sYear, string sSeq)
        {
            Cursor = Cursors.WaitCursor;

            string sAppearYn = RdgbAppearYn.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.ISSUE_YEAR ");
            strSql.AppendLine("      , A.ISSUE_SEQ ");
            strSql.AppendLine("      , A.SEQ ");
            strSql.AppendLine("      , A.OCCUR_DT ");
            strSql.AppendLine("      , A.MFY_ID ");
            strSql.AppendLine("      , A.EMP_ID ");
            strSql.AppendLine("      , A.EMP_NM ");
            strSql.AppendLine("      , A.CHN_NM ");
            strSql.AppendLine("      , A.ISSUE_YMD ");
            strSql.AppendLine("      , A.ENTRANCE_YMD ");
            strSql.AppendLine("      , A.NOW_YMD ");
            strSql.AppendLine("      , A.DEPT_CD ");
            strSql.AppendLine("      , A.JOBPOSITION_CD ");
            strSql.AppendLine("      , A.GRADE_CD ");
            strSql.AppendLine("      , A.IDT_NO ");
            strSql.AppendLine("      , A.IDT_NO AS IDT_NO2 ");
            //strSql.AppendLine("      , CASE WHEN A.IDT_NO IS NULL OR A.IDT_NO = '' THEN ''                                                     ");
            //strSql.AppendLine("             WHEN '" + sAppearYn + "' = 'Y' THEN A.IDT_NO                                                       ");
            //strSql.AppendLine("             ELSE SUBSTRING(A.IDT_NO, 1, 8) + REPLICATE('*', 14 - LEN(SUBSTRING(A.IDT_NO, 1, 8))) END AS IDT_NO ");
            strSql.AppendLine("      , A.ADDR ");
            strSql.AppendLine("      , A.ISSUE_PPS ");
            strSql.AppendLine("   FROM HR_CERT_EMPT_HISTORY A ");
            strSql.AppendLine("  WHERE ISSUE_YEAR = '" + sYear + "' ");
            strSql.AppendLine("    AND ISSUE_SEQ = " + sSeq + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    string sIdtNo = row["IDT_NO"]?.ToString();
                    string sIssuYMd = row["ISSUE_YMD"]?.ToString();

                    if (!string.IsNullOrEmpty(sIssuYMd))
                    {
                        DateTime dtIssuYMD = DateTime.ParseExact(sIssuYMd, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None);
                        DateTime editDate = new DateTime(2023, 02, 07);

                        if (!string.IsNullOrEmpty(sIdtNo) && dtIssuYMD >= editDate)
                        {
                            sIdtNo = ComnEtcFunc.Decrypt(sIdtNo, ComnEtcFunc._SECRET_KEY2);

                            row["IDT_NO2"] = sIdtNo;

                            if (!string.IsNullOrEmpty(sIdtNo) && !string.IsNullOrEmpty(sAppearYn) && sAppearYn.Equals("N"))
                            {
                                row["IDT_NO"] = sIdtNo.Substring(0, 8) + "******";
                            }
                            else
                            {
                                row["IDT_NO"] = sIdtNo;
                            }
                        }
                    }
                }
            }

            GridHis.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void HrCertificateOfEmployment_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void HrCertificateOfEmployment_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
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
                BtnPrint_Click(null, null);
            }
        }

        private void LkupEmpNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void RdgbEmplGB_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);

            DataTable dt = GetLookUpData("3", "A", "A");
            ComGrid.SetLookUpEdit(LkupEmpNm, dt, "CD", "CD", "Y");
            LkupEmpNm.Properties.PopulateColumns();
            LkupEmpNm.Properties.Columns[1].Visible = false;
        }

        private void BGridViewIssue_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.Equals("ISSUE_YMD"))
            {
                string sValue = e.DisplayText;
                if (sValue.Length == 8)
                {
                    string sTemp = sValue.Substring(0, 4) + "-" + sValue.Substring(4, 2) + "-" + sValue.Substring(6, 2);
                    e.DisplayText = sTemp;
                }
            }
            else if (e.Column.FieldName.Equals("ENTRANCE_YMD"))
            {
                string sValue = e.DisplayText;
                if (sValue.Length == 8)
                {
                    string sTemp = sValue.Substring(0, 4) + "-" + sValue.Substring(4, 2) + "-" + sValue.Substring(6, 2);
                    e.DisplayText = sTemp;
                }
            }
            else if (e.Column.FieldName.Equals("NOW_YMD"))
            {
                string sValue = e.DisplayText;
                if (sValue.Length == 8)
                {
                    string sTemp = sValue.Substring(0, 4) + "-" + sValue.Substring(4, 2) + "-" + sValue.Substring(6, 2);
                    e.DisplayText = sTemp;
                }
            }
        }

        private void BGridViewHis_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.Equals("ISSUE_YMD"))
            {
                string sValue = e.DisplayText;
                if (sValue.Length == 8)
                {
                    string sTemp = sValue.Substring(0, 4) + "-" + sValue.Substring(4, 2) + "-" + sValue.Substring(6, 2);
                    e.DisplayText = sTemp;
                }
            }
            else if (e.Column.FieldName.Equals("ENTRANCE_YMD"))
            {
                string sValue = e.DisplayText;
                if (sValue.Length == 8)
                {
                    string sTemp = sValue.Substring(0, 4) + "-" + sValue.Substring(4, 2) + "-" + sValue.Substring(6, 2);
                    e.DisplayText = sTemp;
                }
            }
            else if (e.Column.FieldName.Equals("NOW_YMD"))
            {
                string sValue = e.DisplayText;
                if (sValue.Length == 8)
                {
                    string sTemp = sValue.Substring(0, 4) + "-" + sValue.Substring(4, 2) + "-" + sValue.Substring(6, 2);
                    e.DisplayText = sTemp;
                }
            }
            else if (e.Column.FieldName.Equals("OCCUR_DT"))
            {
                string sValue = e.DisplayText;
                if (sValue.Length == 8)
                {
                    string sTemp = sValue.Substring(0, 4) + "-" + sValue.Substring(4, 2) + "-" + sValue.Substring(6, 2);
                    e.DisplayText = sTemp;
                }
            }
        }

        private void HrCertificateOfEmployment_FormClosed(object sender, FormClosedEventArgs e)
        {
            ComLib.ClsFunc.SaveGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "HrCertificateOfEmployment", BGridViewIssue);
            ComLib.ClsFunc.SaveGridViewLayout(FmMainToolBar2.UserID, "AccAdm", "HrCertificateOfEmployment1", BGridViewHis);
        }

        private void RdgbAppearYn_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sEmpNm = GridViewEmp.GetFocusedRowCellValue(GridColEmpNm)?.ToString() ?? string.Empty;
            BtnRetr.PerformClick();
            GridViewEmp.FocusedRowHandle = GridViewEmp.LocateByDisplayText(0, GridColEmpNm, sEmpNm);

        }

        private void GridViewEmp_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GridViewEmp_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void HrCertificateOfEmployment_TextChanged(object sender, EventArgs e)
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