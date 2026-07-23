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
using System.Diagnostics;
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
*            2. 레이아웃 전체 저장 설정
*/
namespace AccAdm
{
    public partial class UserMgt : DevExpress.XtraEditors.XtraForm
    {
        public UserMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void UserMgt_Load(object sender, EventArgs e)
        {
            ////

            DataTable dtDept = GetLookUpData("1", "", "", "Y");
            DataTable dtJobGrade = GetLookUpData("2", "", "", "Y");
            DataTable dtPgmGroup = GetLookUpData("3", "", "", "Y");
            
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupDept, dtDept, GridUser, GridColUserDeptCd, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupPstn, dtJobGrade, GridUser, GridColUserPstn, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupGroupNm, dtPgmGroup, GridProgram, GridColPgmGroupNm, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(GridLkupPggub, dtPgmGroup, GridProgram, GridColPgmGroupNm, "CD", "NM", "");
            ////
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            arrGrdView = new GridView[] { GridViewUser, GridViewUserDetail, GridViewProgram, GridViewProgramDetail };
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

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout2.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl2.RestoreLayoutFromXml(sFile);
            }

            sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout3.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl3.RestoreLayoutFromXml(sFile);
            }

            BtnRetr.PerformClick();
        }
        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (TabUserMgt.SelectedPageIndex == 0)
            {
                GetGridUser();
                //GetGridUserDetail();
            }
            if (TabUserMgt.SelectedPageIndex == 1)
            {
                GetGridProgram();
                //GetGridProgramDetail();
            }
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam, string sOther)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("1"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , A.DEPT_CD AS SEQ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  WHERE DEPT_CD <> '0000'");
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
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'Z1'");
            }

            if (sOther.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }
            else
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            }


            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void GetGridUser()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.USRCD  ");
            strSql.AppendLine("      , A.USRID  ");
            strSql.AppendLine("      , A.USRNM  ");
            strSql.AppendLine("      , A.PASSWD  ");
            strSql.AppendLine("      , B.DEPT_CD AS DEPTCD  ");
            strSql.AppendLine("      , B.GRADE_CD AS JKWICD  ");
            strSql.AppendLine("      , C.HP_NO AS MOBLNO  ");
            strSql.AppendLine("      , B.EMP_ID AS INSANO ");
            strSql.AppendLine("      , A.USEYN  ");
            strSql.AppendLine("      , A.USLVL  ");
            strSql.AppendLine("      , A.ISPT_YN  ");
            strSql.AppendLine("      , A.RK  ");
            strSql.AppendLine("      , A.CDATE  ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.CUSER AS NUMERIC) IS NULL THEN A.CUSER ELSE DBO.FN_USRNM(A.CUSER) END AS CUSER");
            strSql.AppendLine("      , A.MDATE  ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MUSER AS NUMERIC) IS NULL THEN A.MUSER ELSE DBO.FN_USRNM(A.MUSER) END AS MUSER");
            strSql.AppendLine("      , (SELECT COUNT(B.USE_Y) FROM ZPGMAUT B WHERE USRCD = A.USRCD) AS PGCNT ");
            strSql.AppendLine("   FROM ZUSRLST A  ");
            strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B  ");
            strSql.AppendLine("     ON A.INSANO = B.EMP_ID  ");
            strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_PERSONAL C  ");
            strSql.AppendLine("     ON B.EMP_ID = C.EMP_ID ");
            strSql.AppendLine("  WHERE USEYN = 'Y'  ");
            strSql.AppendLine("  ORDER BY B.DEPT_CD, B.GRADE_CD, A.USRNM  ");
            
            //strSql.AppendLine(" SELECT A1.USRCD, A1.USRID, A1.USRNM, A1.INSANO, B2.DEPT_NM DEPTNM, B3.COM_NM JKWINM,A1.MOBLNO, A1.USLVL, A1.USEYN, A1.RK,A1.CDATE, fn_USRNM(A1.CUSER) CUSER, A1.MDATE, fn_USRNM(A1.MUSER) MUSER,(SELECT COUNT(*) FROM zpgmaut A0 WHERE USE_Y = 'Y' AND A1.USRCD = A0.USRCD)PGCNT ");
            //strSql.AppendLine("   FROM  ZUSRLST A1 LEFT JOIN ACC_DEPT_CD  B2 ON B2.DEPT_CD <> '0000' AND A1.DEPTCD = B2.DEPT_CD  LEFT JOIN com_base_cd  B3 ON B3.CD_GB = 'GRADE_CD' AND A1.JKWICD = B3.COM_CD ");
            //strSql.AppendLine("  WHERE  A1.USRCD<>'00000'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridUser.DataSource = dt;
        }

        private void GetGridProgram()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.PGMID ");
            strSql.AppendLine("      , A.PGMNM ");
            strSql.AppendLine("      , A.SYSGU ");
            strSql.AppendLine("      , A.PGGRP ");
            strSql.AppendLine("      , A.PGGUB ");
            strSql.AppendLine("      , A.PGTAG ");
            strSql.AppendLine("      , A.PGLVL ");
            strSql.AppendLine("      , A.USEYN ");
            strSql.AppendLine("      , A.RK ");
            strSql.AppendLine("      , A.CDATE ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.CUSER AS NUMERIC) IS NULL THEN A.CUSER ELSE DBO.FN_USRNM(A.CUSER) END AS CUSER");
            strSql.AppendLine("      , A.MDATE  ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MUSER AS NUMERIC) IS NULL THEN A.MUSER ELSE DBO.FN_USRNM(A.MUSER) END AS MUSER");
            strSql.AppendLine("      , ( SELECT COUNT(*) ");
            strSql.AppendLine("            FROM ZPGMAUT X1  ");
            strSql.AppendLine("            LEFT JOIN ZUSRLST X2  ");
            strSql.AppendLine("              ON X1.USRCD = X2.USRCD ");
            strSql.AppendLine("           WHERE X1.PGMID = A.PGMID ");
            strSql.AppendLine("             AND X1.USE_Y = 'Y' ");
            strSql.AppendLine("             AND X2.USEYN = 'Y' ) AS USE_PGM_CNT ");
            strSql.AppendLine("   FROM ZPGMLST A ");
            strSql.AppendLine("  WHERE A.USEYN = 'Y' ");
            strSql.AppendLine("  ORDER BY PGGRP ");

            //strSql.AppendLine(" SELECT  A1.PGGRP,B1.COM_NM GRPNM, A1.PGMID, A1.PGMNM, A1.PGLVL, A1.PGGUB,(SELECT COUNT(*) FROM zpgmaut A0 WHERE A0.PGMID = A1.PGMID AND IFNULL(A0.USE_Y, '') = 'Y')pCNT,A1.CDATE, FN_USRNM(A1.CUSER)CUSER, A1.MDATE, FN_USRNM(A1.MUSER)MUSER");
            //strSql.AppendLine("   FROM ZPGMLST A1 ");
            //strSql.AppendLine("   LEFT JOIN ZPGMAUT A2 On A1.PGMID=A2.PGMID      ");
            //strSql.AppendLine("   LEFT JOIN COM_BASE_CD B1 ON B1.CD_GB='Z1' AND A1.PGGRP=B1.COM_CD      ");
            //strSql.AppendLine("  WHERE 1=1");
            //strSql.AppendLine("  GROUP BY PGMNM ");
            //strSql.AppendLine("  ORDER BY PGGRP, PGTAG");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridProgram.DataSource = dt;
        }
        private void GridViewUser_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            GetGridUserDetail();
        }

        private void GridViewProgram_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            GetGridProgramDetail();
        }

        private void GetGridUserDetail()
        {
            GridUserDetail.DataSource = null;

            StringBuilder strSql = new StringBuilder();

            string sUser = string.Empty;

            sUser = GridViewUser.GetFocusedRowCellValue("USRCD").ToString();

            //strSql.Clear();
            //strSql.AppendLine(" ");
            //strSql.AppendLine(" SELECT A.USRCD ");
            //strSql.AppendLine("      , A.PGMID ");
            //strSql.AppendLine("      , C.PGGRP ");
            //strSql.AppendLine("      , C.PGGUB ");
            //strSql.AppendLine("      , C.PGLVL ");
            //strSql.AppendLine("      , C.PGMNM ");
            //strSql.AppendLine("      , A.USE_Y ");
            //strSql.AppendLine("      , A.ADD_Y ");
            //strSql.AppendLine("      , A.UPD_Y ");
            //strSql.AppendLine("      , A.DEL_Y ");
            //strSql.AppendLine("      , A.PRT_Y ");
            //strSql.AppendLine("      , A.XLS_Y ");
            //strSql.AppendLine("      , A.CDATE ");
            //strSql.AppendLine("      , A.CUSER ");
            //strSql.AppendLine("      , A.MDATE ");
            //strSql.AppendLine("      , A.MUSER ");
            //strSql.AppendLine("   FROM ZPGMAUT A ");
            //strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST B ");
            //strSql.AppendLine("     ON B.USRCD = A.USRCD ");
            //strSql.AppendLine("   LEFT OUTER JOIN ZPGMLST C ");
            //strSql.AppendLine("     ON C.PGMID = A.PGMID");
            //strSql.AppendLine("  WHERE A.USRCD = " + sUser + " ");
            //strSql.AppendLine("  ORDER BY C.PGGRP");

            strSql.AppendLine("SELECT A1.*, A2.*,DBO.fn_getreffpf('Z1',PGGRP)GRPNM");
            strSql.AppendLine("  FROM ZPGMLST A1");
            strSql.AppendLine("  LEFT JOIN ZPGMAUT A2 ON A1.PGMID=A2.PGMID AND A2.USRCD = '" + sUser + "'");
            strSql.AppendLine(" WHERE A1.USEYN='Y'");
            strSql.AppendLine(" ORDER BY PGGRP, PGTAG ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridUserDetail.DataSource = dt;
        }

        private void GetGridProgramDetail()
        {
            GridProgramDetail.DataSource = null;
            StringBuilder strSql = new StringBuilder();
            string sPgmId = string.Empty;
            sPgmId = GridViewProgram.GetFocusedRowCellValue("PGMID")?.ToString();
            if (string.IsNullOrEmpty(sPgmId))
            {
                return;
            }
            //strSql.Clear();
            //strSql.AppendLine("SELECT A1.USRCD, A1.USRID, A1.USRNM, B1.DEPT_NM DEPTNM, B2.COM_NM JKWINM,A1.USLVL,A2.USE_Y, A2.ADD_Y, A2.UPD_Y, A2.DEL_Y, A2.XLS_Y");
            //strSql.AppendLine("  FROM ZUSRLST A1 ");
            //strSql.AppendLine(" INNER JOIN ZPGMAUT A2 ON A1.USRCD=A2.USRCD AND A2.PGMID='" + sPgmId + "'");
            //strSql.AppendLine("  LEFT JOIN ACC_DEPT_CD  B1 ON B1.DEPT_CD<>'0000' AND A1.DEPTCD=B1.DEPT_CD");
            //strSql.AppendLine("  LEFT JOIN COM_BASE_CD  B2 ON B2.CD_GB='GRADE_CD' AND A1.JKWICD=B2.COM_CD ");
            //strSql.AppendLine(" WHERE IFNULL(A1.USEYN,'''')='Y' ");
            //strSql.AppendLine(" ORDER BY A1.USRNM ");

            strSql.AppendLine(" SELECT 'SYS001F01' AS PGMID ");
            strSql.AppendLine("      , A.USRCD ");
            strSql.AppendLine("      , A.USRID ");
            strSql.AppendLine("      , A.USRNM ");
            strSql.AppendLine("      , D.DEPT_NM AS DEPTNM ");
            strSql.AppendLine("      , E.COM_NM AS JKWINM  ");
            strSql.AppendLine("      , A.USLVL ");
            strSql.AppendLine("      , ISNULL(B.USE_Y, 'N') AS USE_Y ");
            strSql.AppendLine("      , ISNULL(B.ADD_Y, 'N') AS ADD_Y ");
            strSql.AppendLine("      , ISNULL(B.UPD_Y, 'N') AS UPD_Y ");
            strSql.AppendLine("      , ISNULL(B.DEL_Y, 'N') AS DEL_Y ");
            strSql.AppendLine("      , ISNULL(B.XLS_Y, 'N') AS XLS_Y ");
            strSql.AppendLine("      , ISNULL(B.APV1_Y, 'N') AS APV1_Y ");
            strSql.AppendLine("      , ISNULL(B.APV2_Y, 'N') AS APV2_Y ");
            strSql.AppendLine("      , ISNULL(B.APV3_Y, 'N') AS APV3_Y ");
            strSql.AppendLine("   FROM ZUSRLST A  ");
            strSql.AppendLine("   LEFT OUTER JOIN ZPGMAUT B  ");
            strSql.AppendLine("     ON A.USRCD = B.USRCD ");
            strSql.AppendLine("    AND B.PGMID = '" + sPgmId + "' ");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C   ");
            strSql.AppendLine("     ON A.INSANO = C.EMP_ID   ");
            strSql.AppendLine("   LEFT JOIN ACC_DEPT_CD D    ");
            strSql.AppendLine("     ON C.DEPT_CD = D.DEPT_CD ");
            strSql.AppendLine("   LEFT JOIN COM_BASE_CD E    ");
            strSql.AppendLine("     ON E.CD_GB = 'GRADE_CD'  ");
            strSql.AppendLine("    AND C.GRADE_CD = E.COM_CD ");
            strSql.AppendLine("  WHERE A.USEYN = 'Y' ");
            strSql.AppendLine("  ORDER BY D.DEPT_CD, C.GRADE_CD, A.USRNM ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridProgramDetail.DataSource = dt;
        }

        private void BtnUserSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;
            DataTable dt;
            if (TabUserMgt.SelectedPage == tabNavigationPage1)
                dt = (DataTable)GridUserDetail.DataSource;
            else
                dt = (DataTable)GridProgramDetail.DataSource;

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            //DataTable dtMerge = dt;
            DataTable dtMerge = dsSave.Tables[0];
            if(dtMerge.Rows.Count == 0)
            {
                XtraMessageBox.Show("권한리스트에 수정된 데이터가 존재하지 않습니다.");
                return;
            }
            string sUserCd = string.Empty;
            string sPgMid = string.Empty;
            string sUseY = string.Empty;
            string sAddY = string.Empty;
            string sUpd = string.Empty;
            string sDelY = string.Empty;
            string sXlsY = string.Empty;
            string sApv1Y = string.Empty;
            string sApv2Y = string.Empty;
            string sApv3Y = string.Empty;
            string sCdate = string.Empty;
            string sCUser = string.Empty;
            string sMdate = string.Empty;
            string sMUser = string.Empty;
            
            StringBuilder strSql = new StringBuilder();

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                if (dtMerge.Rows.Count > 0)
                {
                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        sUserCd = GridViewUser.GetFocusedRowCellValue("USRCD").ToString();
                        string sUserNm = GridViewUser.GetFocusedRowCellValue("USRNM")?.ToString();

                        sPgMid = Convert.ToString(dtMerge.Rows[j]["PGMID"]);
                        string sPgmNm = dtMerge.Rows[j]["PGMNM"]?.ToString();

                        sUseY = Convert.ToString(dtMerge.Rows[j]["USE_Y"]);
                        if (string.IsNullOrEmpty(sUseY)) sUseY = "N";

                        sAddY = Convert.ToString(dtMerge.Rows[j]["ADD_Y"]);
                        if (string.IsNullOrEmpty(sAddY)) sAddY = "N";

                        sUpd = Convert.ToString(dtMerge.Rows[j]["UPD_Y"]);
                        if (string.IsNullOrEmpty(sUpd)) sUpd = "N";

                        sDelY = Convert.ToString(dtMerge.Rows[j]["DEL_Y"]);
                        if (string.IsNullOrEmpty(sDelY)) sDelY = "N";

                        sXlsY = Convert.ToString(dtMerge.Rows[j]["XLS_Y"]);
                        if (string.IsNullOrEmpty(sXlsY)) sXlsY = "N";

                        sApv1Y = Convert.ToString(dtMerge.Rows[j]["APV1_Y"]);
                        if (string.IsNullOrEmpty(sApv1Y)) sApv1Y = "N";

                        sApv2Y = Convert.ToString(dtMerge.Rows[j]["APV2_Y"]);
                        if (string.IsNullOrEmpty(sApv2Y)) sApv2Y = "N";

                        sApv3Y = Convert.ToString(dtMerge.Rows[j]["APV3_Y"]);
                        if (string.IsNullOrEmpty(sApv3Y)) sApv3Y = "N";

                        #region LOG
                        ////로그입력 위한 변수세팅
                        //string sPrv_UseY = string.Empty;
                        //string sPrv_AddY = string.Empty;
                        //string sPrv_UpdY = string.Empty;
                        //string sPrv_DelY = string.Empty;
                        //string sPrv_XlsY = string.Empty;
                        //string sPrv_Apv1 = string.Empty;
                        //string sPrv_Apv2 = string.Empty;
                        //string sPrv_Apv3 = string.Empty;
                        //string sLOG = string.Empty;

                        ////이전 데이터정보 조회
                        //strSql.Clear();
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" SELECT * ");
                        //strSql.AppendLine("   FROM ZPGMAUT A ");
                        //strSql.AppendLine("  WHERE A.USRCD = " + sUserCd + "");
                        //strSql.AppendLine("    AND A.PGMID = '" + sPgMid + "'");

                        //DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                        //if (dtPrv.Rows.Count > 0)
                        //{
                        //    sPrv_UseY = dtPrv.Rows[0]["USE_Y"]?.ToString();
                        //    sPrv_AddY = dtPrv.Rows[0]["ADD_Y"]?.ToString();
                        //    sPrv_UpdY = dtPrv.Rows[0]["UPD_Y"]?.ToString();
                        //    sPrv_DelY = dtPrv.Rows[0]["DEL_Y"]?.ToString();
                        //    sPrv_XlsY = dtPrv.Rows[0]["XLS_Y"]?.ToString();
                        //    sPrv_Apv1 = dtPrv.Rows[0]["APV1_Y"]?.ToString();
                        //    sPrv_Apv2 = dtPrv.Rows[0]["APV2_Y"]?.ToString();
                        //    sPrv_Apv3 = dtPrv.Rows[0]["APV3_Y"]?.ToString();

                        //    sLOG = string.Concat("[사용자권한]사용자명 : ", sUserNm,
                        //    ", 프로그램ID : ", sPgmNm,
                        //    ", 사용 : ", sPrv_UseY, " ▶ ", sUseY,
                        //    ", 수정 : ", sPrv_UpdY, " ▶ ", sUpd,
                        //    ", 삭제 : ", sPrv_DelY, " ▶ ", sDelY,
                        //    ", 엑셀 : ", sPrv_XlsY, " ▶ ", sXlsY,
                        //    ", 승인1 : ", sPrv_Apv1, " ▶ ", sApv1Y,
                        //    ", 승인2 : ", sPrv_Apv2, " ▶ ", sApv2Y,
                        //    ", 승인3 : ", sPrv_Apv3, " ▶ ", sApv3Y);
                        //}
                        //else
                        //{
                        //    sLOG = string.Concat("[사용자권한]사용자명 : ", sUserNm,
                        //           ", 프로그램ID : ", sPgmNm,
                        //           ", 사용 : ", sUseY,
                        //           ", 수정 : ", sUpd,
                        //           ", 삭제 : ", sDelY,
                        //           ", 엑셀 : ", sXlsY,
                        //           ", 승인1 : ", sApv1Y,
                        //           ", 승인2 : ", sApv2Y,
                        //           ", 승인3 : ", sApv3Y);
                        //}
                        #endregion

                        strSql.Clear();
                        strSql.AppendLine("IF EXISTS(SELECT* FROM ZPGMAUT WHERE USRCD = "+ sUserCd + " AND PGMID = '"+ sPgMid + "')");
                        strSql.AppendLine("    BEGIN                                                     ");
                        strSql.AppendLine("          UPDATE ZPGMAUT                                      ");
                        strSql.AppendLine("             SET USE_Y = '"+ sUseY + "'                                   ");
		                strSql.AppendLine("               , ADD_Y = '"+ sAddY + "'                                   ");
		                strSql.AppendLine("               , UPD_Y = '"+ sUpd + "'                                   ");
		                strSql.AppendLine("               , DEL_Y = '"+ sDelY + "'                                   ");
		                strSql.AppendLine("               , XLS_Y = '"+ sXlsY + "'                                   ");
		                strSql.AppendLine("               , APV1_Y = '"+ sApv1Y + "'                                  ");
		                strSql.AppendLine("               , APV2_Y = '"+ sApv2Y + "'                                  ");
		                strSql.AppendLine("               , APV3_Y = '"+ sApv3Y + "'                                  ");
		                strSql.AppendLine("               , MUSER = "+ FmMainToolBar2.UserID + "                                   ");
		                strSql.AppendLine("               , MDATE = GETDATE()                            ");
                        strSql.AppendLine("           WHERE USRCD = "+ sUserCd + "                                    ");
                        strSql.AppendLine("             AND PGMID = '"+ sPgMid + "'                                   ");
                        strSql.AppendLine("      END                                                     ");
                        strSql.AppendLine(" ELSE                                                         ");
                        strSql.AppendLine("    BEGIN                                                     ");
                        strSql.AppendLine("           INSERT INTO ZPGMAUT                                ");
                        strSql.AppendLine("                (USRCD, PGMID                                 ");
                        strSql.AppendLine("                , USE_Y, ADD_Y, UPD_Y , DEL_Y, XLS_Y          ");
                        strSql.AppendLine("                , APV1_Y, APV2_Y, APV3_Y, CUSER, CDATE )      ");
                        strSql.AppendLine("          VALUES("+ sUserCd + ", '"+ sPgMid + "'                                        ");
                        strSql.AppendLine("                , '"+ sUseY + "', '"+ sAddY + "', '"+ sUpd + "', '"+ sDelY + "', '"+ sXlsY + "'");
                        strSql.AppendLine("                , '"+ sApv1Y + "', '"+ sApv2Y + "', '"+ sApv3Y + "', "+ FmMainToolBar2.UserID + ", GETDATE())");
                        strSql.AppendLine("      END                                                     ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    GetGridUserDetail();
                }
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
            Cursor = Cursors.Default;

        }
        private void BtnProgramSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            Cursor = Cursors.WaitCursor;
            DataTable dt = (DataTable)GridProgramDetail.DataSource;

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMerge = dsSave.Tables[0];
            string sUserCd = string.Empty;
            string sPgMid = string.Empty;
            string sUseY = string.Empty;
            string sAddY = string.Empty;
            string sUpd = string.Empty;
            string sDelY = string.Empty;
            string sXlsY = string.Empty;
            string sCUser = string.Empty;
            string sMUser = string.Empty;
            
            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            StringBuilder strSql = new StringBuilder();

            try
            {
                if (dtMerge.Rows.Count > 0)
                {
                    for (int j = 0; j < dtMerge.Rows.Count; j++)
                    {
                        sUserCd = dtMerge.Rows[j]["USRCD"]?.ToString();
                        string sUserNm = dtMerge.Rows[j]["USRNM"]?.ToString();

                        sPgMid = GridViewProgram.GetFocusedRowCellValue("PGMID").ToString();
                        string sPgmNm = GridViewProgram.GetFocusedRowCellValue("PGMNM")?.ToString();
                        
                        sUseY = Convert.ToString(dtMerge.Rows[j]["USE_Y"]);
                        if (string.IsNullOrEmpty(sUseY)) sUseY = "N";

                        sAddY = Convert.ToString(dtMerge.Rows[j]["ADD_Y"]);
                        if (string.IsNullOrEmpty(sAddY)) sAddY = "N";

                        sUpd = Convert.ToString(dtMerge.Rows[j]["UPD_Y"]);
                        if (string.IsNullOrEmpty(sUpd)) sUpd = "N";

                        sDelY = Convert.ToString(dtMerge.Rows[j]["DEL_Y"]);
                        if (string.IsNullOrEmpty(sDelY)) sDelY = "N";

                        sXlsY = Convert.ToString(dtMerge.Rows[j]["XLS_Y"]);
                        if (string.IsNullOrEmpty(sXlsY)) sXlsY = "N";

                        #region LOG
                        ////로그입력 위한 변수세팅
                        //string sPrv_UseY = string.Empty;
                        //string sPrv_AddY = string.Empty;
                        //string sPrv_UpdY = string.Empty;
                        //string sPrv_DelY = string.Empty;
                        //string sPrv_XlsY = string.Empty;
                        //string sPrv_Apv1 = string.Empty;
                        //string sPrv_Apv2 = string.Empty;
                        //string sPrv_Apv3 = string.Empty;
                        //string sLOG = string.Empty;

                        ////이전 데이터정보 조회
                        //strSql.Clear();
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" SELECT * ");
                        //strSql.AppendLine("   FROM ZPGMAUT A ");
                        //strSql.AppendLine("  WHERE A.USRCD = " + sUserCd + "");
                        //strSql.AppendLine("    AND A.PGMID = '" + sPgMid + "'");

                        //DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                        //if (dtPrv.Rows.Count > 0)
                        //{
                        //    sPrv_UseY = dtPrv.Rows[0]["USE_Y"]?.ToString();
                        //    sPrv_AddY = dtPrv.Rows[0]["ADD_Y"]?.ToString();
                        //    sPrv_UpdY = dtPrv.Rows[0]["UPD_Y"]?.ToString();
                        //    sPrv_DelY = dtPrv.Rows[0]["DEL_Y"]?.ToString();
                        //    sPrv_XlsY = dtPrv.Rows[0]["XLS_Y"]?.ToString();
                        //    sPrv_Apv1 = dtPrv.Rows[0]["APV1_Y"]?.ToString();
                        //    sPrv_Apv2 = dtPrv.Rows[0]["APV2_Y"]?.ToString();
                        //    sPrv_Apv3 = dtPrv.Rows[0]["APV3_Y"]?.ToString();

                        //    sLOG = string.Concat("[사용자권한]사용자명 : ", sUserNm,
                        //    ", 프로그램ID : ", sPgmNm,
                        //    ", 사용 : ", sPrv_UseY, " ▶ ", sUseY,
                        //    ", 수정 : ", sPrv_UpdY, " ▶ ", sUpd,
                        //    ", 삭제 : ", sPrv_DelY, " ▶ ", sDelY,
                        //    ", 엑셀 : ", sPrv_XlsY, " ▶ ", sXlsY);
                        //}
                        //else
                        //{
                        //    sLOG = string.Concat("[사용자권한]사용자명 : ", sUserNm,
                        //           ", 프로그램ID : ", sPgmNm,
                        //           ", 사용 : ", sUseY,
                        //           ", 수정 : ", sUpd,
                        //           ", 삭제 : ", sDelY,
                        //           ", 엑셀 : ", sXlsY);
                        //}
                        #endregion

                        strSql.Clear();
                        strSql.AppendLine("IF EXISTS(SELECT* FROM ZPGMAUT WHERE USRCD = " + sUserCd + " AND PGMID = '" + sPgMid + "')");
                        strSql.AppendLine("    BEGIN                                                     ");
                        strSql.AppendLine("          UPDATE ZPGMAUT                                      ");
                        strSql.AppendLine("             SET USE_Y = '" + sUseY + "'                                   ");
                        strSql.AppendLine("               , ADD_Y = '" + sAddY + "'                                   ");
                        strSql.AppendLine("               , UPD_Y = '" + sUpd + "'                                   ");
                        strSql.AppendLine("               , DEL_Y = '" + sDelY + "'                                   ");
                        strSql.AppendLine("               , XLS_Y = '" + sXlsY + "'                                   ");
                        strSql.AppendLine("               , MUSER = " + FmMainToolBar2.UserID + "                                   ");
                        strSql.AppendLine("               , MDATE = GETDATE()                            ");
                        strSql.AppendLine("           WHERE USRCD = " + sUserCd + "                                    ");
                        strSql.AppendLine("             AND PGMID = '" + sPgMid + "'                                   ");
                        strSql.AppendLine("      END                                                     ");
                        strSql.AppendLine(" ELSE                                                         ");
                        strSql.AppendLine("    BEGIN                                                     ");
                        strSql.AppendLine("           INSERT INTO ZPGMAUT                                ");
                        strSql.AppendLine("                (USRCD, PGMID                                 ");
                        strSql.AppendLine("                , USE_Y, ADD_Y, UPD_Y , DEL_Y, XLS_Y          ");
                        strSql.AppendLine("                , CUSER, CDATE )      ");
                        strSql.AppendLine("          VALUES(" + sUserCd + ", '" + sPgMid + "'                                        ");
                        strSql.AppendLine("                , '" + sUseY + "', '" + sAddY + "', '" + sUpd + "', '" + sDelY + "', '" + sXlsY + "'");
                        strSql.AppendLine("                ,  " + FmMainToolBar2.UserID + ", GETDATE())");
                        strSql.AppendLine("      END                                                     ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");
                    GetGridProgramDetail();
                }
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
            Cursor = Cursors.Default;
        }
        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        public string _USRID = string.Empty;
        public string _PRGID = string.Empty;
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (TabUserMgt.SelectedPageIndex == 0)
            {
                PopUpUserInsert fm = new PopUpUserInsert();
                fm.Owner = this;
                if(fm.ShowDialog() == DialogResult.OK)
                {
                    GetGridUser();
                    GridViewUser.FocusedRowHandle = GridViewUser.LocateByDisplayText(0, GridColUsrId, _USRID);
                    GetGridUserDetail();

                    _USRID = string.Empty;
                }
            }
            if (TabUserMgt.SelectedPageIndex == 1)
            {
                PopUpProgramInsert fm = new PopUpProgramInsert();
                fm.Owner = this;
                if(fm.ShowDialog() == DialogResult.OK)
                {
                    GetGridProgram();
                    GridViewProgram.FocusedRowHandle = GridViewProgram.LocateByDisplayText(0, GridColPGMID, _PRGID);
                    GetGridProgramDetail();

                    _PRGID = string.Empty;
                }
            }
        }

        private void GridViewUser_DoubleClick(object sender, EventArgs e)
        {
            string s1 = GridViewUser.GetFocusedRowCellValue("USRCD").ToString();

            PopUpUserInsert obj = new PopUpUserInsert();
            //StringBuilder strSql = new StringBuilder();

            //strSql.Clear();
            //strSql.AppendLine(" SELECT USRCD, USRID, USRNM, PASSWD, DEPTCD, JKWICD, MOBLNO, INSANO, USEYN, USLVL, ISPT_YN, RK, CDATE, CUSER, MDATE, MUSER");
            //strSql.AppendLine("   FROM  ZUSRLST A1 ");
            //strSql.AppendLine("  WHERE  USRCD ='" + s1 + "'");

            //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            //if (dt.Rows.Count == 0) return;

            obj.drUserInfo = GridViewUser.GetFocusedDataRow();
            obj.Owner = this;
            if (obj.ShowDialog() == DialogResult.OK)
            {
                GetGridUser();

                GridViewUser.FocusedRowHandle = GridViewUser.LocateByDisplayText(0, GridColUsrId, _USRID);
                _USRID = string.Empty;
            }

        }

        private void GridViewProgram_DoubleClick(object sender, EventArgs e)
        {
            string s1 = GridViewProgram.GetFocusedRowCellValue("PGMNM").ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT PGMID, PGMNM, SYSGU, PGGRP, PGGUB, PGTAG, PGLVL, USEYN, RK, CDATE, CUSER, MDATE, MUSER");
            strSql.AppendLine("   FROM  ZPGMLST A1 ");
            strSql.AppendLine("  WHERE  PGMNM ='" + s1 + "'");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if(dt.Rows.Count == 0)
            {
                XtraMessageBox.Show("해당 프로그램 정보가 없습니다");
                return;
            }

            string rsPgmNm = dt.Rows[0]["PGMNM"].ToString();
            string rsPgmId = dt.Rows[0]["PGMID"].ToString();
            string rsPggrp = dt.Rows[0]["PGGRP"].ToString();
            string rsPggub = dt.Rows[0]["PGGUB"].ToString();
            string rsPglvl = dt.Rows[0]["PGLVL"].ToString();
            string rsPgTag = dt.Rows[0]["PGTAG"].ToString();
            string rsRK =    dt.Rows[0]["RK"].ToString();
            string rsUseYn = dt.Rows[0]["USEYN"].ToString();

            PopUpProgramInsert obj = new PopUpProgramInsert();

            obj.drPgm = dt.Rows[0];
            obj.Owner = this;
            if(obj.ShowDialog() == DialogResult.OK)
            {
                GetGridProgram();

                GridViewProgram.FocusedRowHandle = GridViewProgram.LocateByDisplayText(0, GridColPGMID, _PRGID);
                _PRGID = string.Empty;
            }

        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            int iIDX = 0;
            
            if(TabUserMgt.SelectedPageIndex == 0)
            {
                iIDX = GridViewUser.FocusedRowHandle;
            }
            else
            {
                iIDX = GridViewProgram.FocusedRowHandle;
            }

            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                if (TabUserMgt.SelectedPageIndex == 0)
                {
                    if (GridViewUser.GetFocusedRowCellValue("USRCD") == null) return;
                    if ((MessageBox.Show(this, "삭제 하시겠습니까???", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No))
                    {
                        return;
                    }

                    string sUserCd = GridViewUser.GetFocusedRowCellValue("USRCD").ToString();
                    string sUserNM = GridViewUser.GetFocusedRowCellValue("USRNM").ToString();

                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    strSql.AppendLine(" UPDATE ZUSRLST ");
                    strSql.AppendLine("    SET USEYN = 'N' ");
                    strSql.AppendLine("  WHERE  USRCD ='" + sUserCd + "'");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    string sLogRmk = string.Concat("[사용자관리]사용자코드 : ", sUserCd, ", 유저명 : ", sUserNM);

                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Add("CUR_TIME", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    dicParams.Add("USRCD", FmMainToolBar2.UserID);
                    dicParams.Add("PGMID", this.Name);
                    dicParams.Add("IP", ComnEtcFunc.GetLocalIP());
                    dicParams.Add("EDIT_RMK", sLogRmk);

                    strSql.Clear();
                    strSql.AppendLine(" INSERT INTO ZSYS_LOG ");
                    strSql.AppendLine(" 	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                    strSql.AppendLine(" 	      VALUES ");
                    strSql.AppendLine(" 	           ( @CUR_TIME ");
                    strSql.AppendLine(" 	           , @USRCD ");
                    strSql.AppendLine(" 	           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    strSql.AppendLine("                      FROM ZSYS_LOG X1 ");
                    strSql.AppendLine("                     WHERE X1.OCCUR_DATE = @CUR_TIME ");
                    strSql.AppendLine("                       AND X1.USRCD = @USRCD ) --LOG_SEQ(구분자) ");
                    strSql.AppendLine(" 	           , @PGMID ");
                    strSql.AppendLine(" 	           , 'D' ");
                    strSql.AppendLine(" 	           , @IP ");
                    strSql.AppendLine(" 	           , @EDIT_RMK ); ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    foreach (KeyValuePair<string, string> param in dicParams)
                    {
                        cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                    
                }
                else if (TabUserMgt.SelectedPageIndex == 1)
                {
                    if (GridViewProgram.GetFocusedRowCellValue("PGMID") == null) return;
                    if ((MessageBox.Show(this, "삭제 하시겠습니까???", "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No))
                    {
                        return;
                    }

                    string sPgmId = GridViewProgram.GetFocusedRowCellValue("PGMID").ToString();
                    string sPgmNm = GridViewProgram.GetFocusedRowCellValue(GridColPgmNm).ToString();

                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    strSql.AppendLine(" DELETE ");
                    strSql.AppendLine("   FROM  ZPGMLST  ");
                    strSql.AppendLine("  WHERE  PGMID ='" + sPgmId + "'");
                    
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    string sLogRmk = string.Concat("[프로그램관리]프로그램ID : ", sPgmId, ", 프로그램명 : ", sPgmNm);

                    Dictionary<string, string> dicParams = new Dictionary<string, string>();

                    dicParams.Add("CUR_TIME", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    dicParams.Add("USRCD", FmMainToolBar2.UserID);
                    dicParams.Add("PGMID", this.Name);
                    dicParams.Add("IP", ComnEtcFunc.GetLocalIP());
                    dicParams.Add("EDIT_RMK", sLogRmk);

                    strSql.Clear();
                    strSql.AppendLine(" INSERT INTO ZSYS_LOG ");
                    strSql.AppendLine(" 	           ( OCCUR_DATE, USRCD, LOG_SEQ, PGM_ID , EDIT_KIND, ACS_IP, EDIT_RMK )   ");
                    strSql.AppendLine(" 	      VALUES ");
                    strSql.AppendLine(" 	           ( @CUR_TIME ");
                    strSql.AppendLine(" 	           , @USRCD ");
                    strSql.AppendLine(" 	           , ( SELECT ISNULL(MAX(X1.LOG_SEQ), 0) + 1  ");
                    strSql.AppendLine("                      FROM ZSYS_LOG X1 ");
                    strSql.AppendLine("                     WHERE X1.OCCUR_DATE = @CUR_TIME ");
                    strSql.AppendLine("                       AND X1.USRCD = @USRCD ) --LOG_SEQ(구분자) ");
                    strSql.AppendLine(" 	           , @PGMID ");
                    strSql.AppendLine(" 	           , 'D' ");
                    strSql.AppendLine(" 	           , @IP ");
                    strSql.AppendLine(" 	           , @EDIT_RMK ); ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    foreach (KeyValuePair<string, string> param in dicParams)
                    {
                        cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                    }
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" DELETE ");
                    strSql.AppendLine("   FROM ZPGMAUT ");
                    strSql.AppendLine("  WHERE PGMID = '" + sPgmId + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제를 완료했습니다.");

                BtnRetr.PerformClick();

                if(TabUserMgt.SelectedPageIndex == 0)
                {
                    GridViewUser.FocusedRowHandle = iIDX - 1;
                }
                else
                {
                    GridViewProgram.FocusedRowHandle = iIDX - 1;
                }

                Cursor = Cursors.Default;
            }
            catch(Exception ex)
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewUserDetail_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {

            if (e.Column.FieldName == "USE_Y")
            {

                if ((sender as GridView).GetRowCellValue(e.RowHandle, "USE_Y").ToString().Equals("Y"))
                {
                    e.Appearance.BackColor = Color.Yellow;
                }

            }
            if (e.Column.FieldName == "ADD_Y")
            {
                if ((sender as GridView).GetRowCellValue(e.RowHandle, "ADD_Y").ToString().Equals("Y"))
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
            }
            if (e.Column.FieldName == "UPD_Y")
            {
                if ((sender as GridView).GetRowCellValue(e.RowHandle, "UPD_Y").ToString().Equals("Y"))
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
            }
            if (e.Column.FieldName == "DEL_Y")
            {
                if ((sender as GridView).GetRowCellValue(e.RowHandle, "DEL_Y").ToString().Equals("Y"))
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
            }
            if (e.Column.FieldName == "XLS_Y")
            {
                if ((sender as GridView).GetRowCellValue(e.RowHandle, "XLS_Y").ToString().Equals("Y"))
                {
                    e.Appearance.BackColor = Color.Yellow;
                }
            }

        }

        private void GridViewProgramDetail_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridViewUserDetail_RowCellStyle(sender, e);
            
        }

        private void TabUserMgt_SelectedPageChanged(object sender, DevExpress.XtraBars.Navigation.SelectedPageChangedEventArgs e)
        {
            if(TabUserMgt.SelectedPageIndex == 0)
            {
                GetGridUser();
                GetGridUserDetail();               
            }
            
            if(TabUserMgt.SelectedPageIndex == 1 )
            {
                GetGridProgram();
                GetGridProgramDetail();
            }
            
           
        }

        private void UserMgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnDelete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel_Click(null, null);
            }
        }

        private void GridViewUser_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void UserMgt_FormClosed(object sender, FormClosedEventArgs e)
        {

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
                string sFileNM = string.Empty;

                if(TabUserMgt.SelectedPageIndex == 0)
                {
                    sFileNM = "사용자현황 ";
                }
                else if (TabUserMgt.SelectedPageIndex == 1)
                {
                    sFileNM = "프로그램현황 ";
                }
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    if (TabUserMgt.SelectedPageIndex == 0)
                    {
                        GridUser.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
                    else if (TabUserMgt.SelectedPageIndex == 1)
                    {
                        GridProgram.ExportToXls(FileName);
                        Process.Start(FileName);
                    }
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

        private void UserMgt_TextChanged(object sender, EventArgs e)
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
                layoutControl2.SaveLayoutToXml(path + @"\" + this.Name + "_Layout2.xaml");
                layoutControl3.SaveLayoutToXml(path + @"\" + this.Name + "_Layout3.xaml");

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }
    }
}