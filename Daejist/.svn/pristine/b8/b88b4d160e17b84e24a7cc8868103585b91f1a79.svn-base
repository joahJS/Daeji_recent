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
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using ComLib;
using System.Data.SqlClient;
using DevExpress.XtraSplashScreen;
using DevExpress.DataAccess.Excel;
using System.Collections;
using MySql.Data.MySqlClient;

/*
 * 작성자: 정은영
 * 작성일: 모름
 * 
 * ------------------------------------
 * 
 * 수정일자:2022-12-16
 * 수정자:정은영
 * ID:#0001
 * 내용:(현업요청)
 *      1. 지각,외출은 0~30일때 0.5, 30~60일때 1.0 으로 계산
 */

namespace AccAdm
{
    public partial class GE001F00 : DevExpress.XtraEditors.XtraForm
    {
        public GE001F00()
        {
            InitializeComponent();
        }

        public string _BASDT = string.Empty;
        public string _EMPID = string.Empty;

        private static double _YAGAN = 18; //야간 기준시간. 18시 이후 출근자: 야간
        private static string _NTBAB = string.Empty;//잔업시작시간

        private static string _YAGAN1 = string.Empty;//야간시작
        private static string _YAGAN2 = string.Empty;//야간종료

        private static string _ILBAN1 = string.Empty;//일반출근
        private static string _ILBAN2 = string.Empty;//일반퇴근
        private static string _YUKIC1 = string.Empty;//여직원출근
        private static string _YUKIC2 = string.Empty;//여직원퇴근

        private void GE001F00_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();
            SetGETimes();

            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.SetDateToValue(DateFrom);

            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            dicParams.Clear();
            dicParams.Add("CMD", "COMCD");
            dicParams.Add("RCDTP", "WKGUB");//근태구분
            DataTable dtWkgub = DBConn.GetDataTable(DBConn.dbCon, "DP_GET_REFNO", dicParams);
            ComLib.ComGrid.SetGridLookUpEdit(RepoGridLkupWkgub, dtWkgub, GridRetr, GridColWkGub, "CD", "NM", "");

            //사원명
            ComnEtcFunc.SetBoundGridLookUp(RepoGridLkupEmpid, "HR_EMP_BASIS", "", "");
            //부서명
            ComnEtcFunc.SetBoundGridLookUp(RepoGridLkupDept, "ACC_DEPT_CD", "DEPT_CD", "DEPT_NM");

            BtnRetr.PerformClick();
        }

        private void SetGETimes()
        {
            try
            {
                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "COMCD3");
                dicParams.Add("RCDTP", "GMUGB");//근태구분
                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, "DP_GET_REFNO", dicParams);

                if (dt != null)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sComcd = dt.Rows[i]["COM_CD"]?.ToString();
                        string sComSub1 = dt.Rows[i]["COM_SUB_CD1"]?.ToString();
                        string sComSub2 = dt.Rows[i]["COM_SUB_CD2"]?.ToString();

                        if (!string.IsNullOrEmpty(sComcd))
                        {
                            switch (sComcd)
                            {
                                case "0":
                                    _ILBAN1 = sComSub1;//일반출근
                                    _ILBAN2 = sComSub2;//일반퇴근
                                    break;
                                case "1":
                                    _YUKIC1 = sComSub1;//여직원출근
                                    _YUKIC2 = sComSub2;//여직원퇴근
                                    break;
                                case "2":
                                    if (DateTime.TryParse(sComSub1, out DateTime dateTime))
                                    {
                                        _YAGAN = dateTime.Hour; //야간 기준시간. 18시 이후 출근자: 야간
                                    }
                                    break;
                                case "3":
                                    _YAGAN1 = sComSub1;//야간시작
                                    _YAGAN2 = sComSub2;//야간종료
                                    break;
                                case "4":
                                    _NTBAB = sComSub1;//잔업시작시간
                                    break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
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

        private void GE001F00_Shown(object sender, EventArgs e)
        {
            DateFrom.Focus();
        }

        #region 조회
        public void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_BASDT))
            {
                DateFrom.EditValue = DateTime.Parse(_BASDT);
                _BASDT = string.Empty;
            }

            string sDate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string sNmk = TxtNmk.EditValue?.ToString();

            try
            {

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" WITH TEMP1 AS(                                                  ");
                strSql.AppendLine("     --출근                                                      ");
                strSql.AppendLine("     SELECT CONVERT(DATE, E_DATE) AS BASDT                       ");
                strSql.AppendLine("          , ROW_NUMBER() OVER(PARTITION BY E_IDNO ORDER BY E_TIME) AS SEQ");
                strSql.AppendLine("          , E_ID                                                 ");
                strSql.AppendLine("          , E_IDNO                                               ");
                strSql.AppendLine("          , STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS ETIME  ");
                strSql.AppendLine("          , E_MODE                                               ");
                strSql.AppendLine("       FROM TENTER                                               ");
                strSql.AppendLine("      WHERE E_DATE = '" + sDate.Replace("-","") + "'             ");
                strSql.AppendLine("        AND E_MODE = '1'                                         ");
                strSql.AppendLine(" ), TEMP2 AS(                                                    ");
                strSql.AppendLine("     --퇴근                                                      ");
                strSql.AppendLine("     SELECT CONVERT(DATE, E_DATE) AS BASDT                       ");
                strSql.AppendLine("          , ROW_NUMBER() OVER(PARTITION BY E_IDNO ORDER BY E_TIME) AS SEQ");
                strSql.AppendLine("          , E_ID                                                 ");
                strSql.AppendLine("          , E_IDNO                                               ");
                strSql.AppendLine("          , STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS ETIME  ");
                strSql.AppendLine("          , E_MODE                                               ");
                strSql.AppendLine("       FROM TENTER                                               ");
                strSql.AppendLine("      WHERE E_DATE = '" + sDate.Replace("-", "") + "'            ");
                strSql.AppendLine("        AND E_MODE = '2'                                         ");
                strSql.AppendLine(" )                                                               ");
                strSql.AppendLine("                                                                 ");
                strSql.AppendLine(" SELECT 'MODI' AS EDIT                                                ");
                strSql.AppendLine("      , A1.BASDT                                                 ");
                strSql.AppendLine("      , A1.EMPID                                                 ");
                strSql.AppendLine("      , A1.SEQ");
                strSql.AppendLine("      , A2.DEPT_CD");
                strSql.AppendLine("      , A3.DEPT_NM AS DEPTNM");
                strSql.AppendLine("      , A2.EMP_NM                                                ");
                strSql.AppendLine("      , A1.WKGUB                                                 ");
                strSql.AppendLine("      , T1.ETIME AS SCMTM1                                       ");
                strSql.AppendLine("      , T2.ETIME AS SCMTM2                                       ");
                strSql.AppendLine("      , A1.WKINTM                                                ");
                strSql.AppendLine("      , A1.WKOTTM                                                ");
                strSql.AppendLine("      , A1.GOINTM                                                ");
                strSql.AppendLine("      , A1.GOOTTM                                                ");
                strSql.AppendLine("      , A1.WKBASE AS WKBASE                                      ");
                strSql.AppendLine("      , A1.GWKTM1                                                ");
                strSql.AppendLine("      , A1.GWKTM2                                                ");
                strSql.AppendLine("      , A1.GWKTM3                                                ");
                strSql.AppendLine("      , A1.GWKTM4                                                ");
                strSql.AppendLine("      , A1.GWKTM5                                                ");
                strSql.AppendLine("      , A1.GWKTM6                                                ");
                strSql.AppendLine("      , A1.GWKTM7                                                ");
                strSql.AppendLine("      , A1.GWKTM8                                                ");
                strSql.AppendLine("      , A1.GWKTM9                                                ");
                strSql.AppendLine("      , A2.GENDER_CD");
                strSql.AppendLine("      , A1.RK1                                                   ");
                strSql.AppendLine("      , A1.MCHKYN                                                ");
                strSql.AppendLine("      , A1.CDATE                                                 ");
                strSql.AppendLine("      , CASE WHEN TRY_PARSE(A1.CUSER AS NUMERIC) IS NULL THEN A1.CUSER ELSE DBO.FN_USRNM(A1.CUSER) END AS CUSER");
                strSql.AppendLine("      , A1.MDATE");
                strSql.AppendLine("      , CASE WHEN TRY_PARSE(A1.MUSER AS NUMERIC) IS NULL THEN A1.MUSER ELSE DBO.FN_USRNM(A1.MUSER) END AS MUSER");
                strSql.AppendLine("   FROM GDAYF A1              ");
                strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS A2  ");
                strSql.AppendLine("     ON A1.EMPID = A2.EMP_ID  ");
                strSql.AppendLine("   LEFT JOIN ACC_DEPT_CD A3       ");
                strSql.AppendLine("        ON A2.DEPT_CD = A3.DEPT_CD");
                strSql.AppendLine("   LEFT JOIN TEMP1 T1         ");
                strSql.AppendLine("     ON A1.EMPID = T1.E_IDNO  ");
                strSql.AppendLine("    AND A1.SEQ = T1.SEQ  ");
                strSql.AppendLine("   LEFT JOIN TEMP2 T2         ");
                strSql.AppendLine("     ON A1.EMPID = T2.E_IDNO  ");
                strSql.AppendLine("    AND A1.SEQ = T2.SEQ  ");
                strSql.AppendLine("  WHERE A1.BASDT = '"+ sDate + "'        ");
                if(!string.IsNullOrEmpty(sNmk))
                    strSql.AppendLine("    AND A2.EMP_NM LIKE '%"+ sNmk + "%'  ");
                strSql.AppendLine("  ORDER BY A2.DEPT_CD, A2.GRADE_CD, A2.EMP_NM");

               DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt != null)
                {
                    GridRetr.DataSource = dt;
                    if (dt.Rows.Count > 0)
                    {
                        GridRetr.Focus();

                        if (!string.IsNullOrEmpty(_EMPID))
                        {
                            GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColEmpid, _EMPID);
                        }

                        SetMagamBtn();
                    }
                    else
                    {
                        DateFrom.Focus();
                    }

                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString(), "조회 오류");
            }

            //연차(공용) 자동사용
            SaveYnGong();
        }

        //연차(공용) 자동사용
        private void SaveYnGong()
        {
            try
            {
                DateTime firstDay = DateTime.Parse(DateFrom.EditValue?.ToString().Substring(0, 7) + "-01");
                DateTime lastDay = firstDay.AddMonths(1).AddDays(-1);

                DateTime dateTime = firstDay;

                StringBuilder strSql = new StringBuilder();

                while (true)
                {
                    if (dateTime > lastDay)
                        break;

                    int weekNum = ComnEtcFunc.GetWeekOfMonth(dateTime);

                    if(weekNum == 2 || weekNum == 4)
                    {
                        if (dateTime.DayOfWeek == DayOfWeek.Saturday)
                        {
                            string sBASDT = dateTime.ToString("yyyy-MM-dd");

                            strSql.Clear();
                            strSql.AppendLine(" SELECT COUNT(*) AS CNT     ");
                            strSql.AppendLine("   FROM GDAYF               ");
                            strSql.AppendLine("  WHERE BASDT = '" + sBASDT + "'");

                            DataTable dtCnt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                            if(dtCnt != null && dtCnt.Rows.Count > 0)
                            {
                                if(int.TryParse(dtCnt.Rows[0]["CNT"]?.ToString(), out int iResult))
                                {
                                    if(iResult > 0)
                                    {
                                        dateTime = dateTime.AddDays(1);
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                dateTime = dateTime.AddDays(1);
                                continue;
                            }

                            strSql.Clear();
                            strSql.AppendLine(" SELECT A1.EMP_ID             ");
                            strSql.AppendLine("  FROM HR_EMP_BASIS A1        ");
                            strSql.AppendLine("  LEFT JOIN GDAYF A2          ");
                            strSql.AppendLine("    ON A1.EMP_ID = A2.EMPID   ");
                            strSql.AppendLine("   AND A2.BASDT = '" + sBASDT + "'");
                            strSql.AppendLine(" WHERE A1.EMPL_GB = 'Y'       ");
                            strSql.AppendLine("   AND A1.GDAYYN = 'Y'        ");
                            strSql.AppendLine("   AND A2.BASDT IS NULL       ");

                            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                            if(dt != null && dt.Rows.Count > 0)
                            {
                                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                                cmd.Transaction = DBConn.dbTran;

                                for (int i=0;i<dt.Rows.Count; i++)
                                {
                                    string sEMPID = dt.Rows[i]["EMP_ID"]?.ToString();
                                    string sRK = "연차(공용)";
                                    string sUSRCD = FmMainToolBar2.drUser["USRCD"]?.ToString();

                                    strSql.Clear();
                                    strSql.AppendLine("INSERT INTO GDAYF( BASDT   ");
                                    strSql.AppendLine("                 , EMPID   ");
                                    strSql.AppendLine("                 , SEQ");
                                    strSql.AppendLine("                 , RK1");
                                    strSql.AppendLine("                 , CUSER)  ");
                                    strSql.AppendLine("           VALUES( '" + sBASDT + "'      ");
                                    strSql.AppendLine("                 , '" + sEMPID + "'");
                                    strSql.AppendLine("                 , 1");
                                    strSql.AppendLine("                 , '" + sRK + "'");
                                    strSql.AppendLine("                 , '" + sUSRCD + "')     ");

                                    cmd.CommandType = CommandType.Text;
                                    cmd.CommandText = strSql.ToString();
                                    cmd.ExecuteNonQuery();
                                }

                                DBConn.dbTran.Commit();
                                DBConn.dbTran = null;
                            }
                        }
                    }

                    dateTime = dateTime.AddDays(1);
                }


            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void SetMagamBtn()
        {
            string sMCHKYN = GridViewRetr.GetFocusedRowCellValue("MCHKYN").ToString();

            if (sMCHKYN == null)
                return;

            if (sMCHKYN.Equals("Y"))
            {
                BtnMagam.Text = "마감취소";
                GridViewRetr.OptionsBehavior.Editable = false;
            }
            else
            {
                BtnMagam.Text = "마감";
                GridViewRetr.OptionsBehavior.Editable = true;
            }
        }

        #endregion

        #region 추가
        private void BtnAdd_Click(object sender, EventArgs e)
        {
            //GE001F01 frm = new GE001F01();

            //frm.Owner = this;
            //if (frm.ShowDialog() == DialogResult.OK)
            //{

            //}

            string sDate = DateFrom.EditValue?.ToString().Substring(0, 10);

            if (ChkMagamYN(sDate))
            {
                XtraMessageBox.Show("마감된 자료는 근태추가가 불가능합니다.");
                return;
            }

            if (ChkPayApply(sDate))
            {
                XtraMessageBox.Show("급여 승인이 완료되어 근태추가를 할 수 없습니다.");
                return;
            }

            if (GridViewRetr.RowCount == 0)
            {
                GridViewRetr.AddNewRow();
                GridViewRetr.UpdateCurrentRow();

                GridViewRetr.SetFocusedRowCellValue("BASDT", DateFrom.EditValue?.ToString().Substring(0, 10));
                GridViewRetr.FocusedColumn = GridColEmpid;
            }
            else
            {
                string sEmpid = GridViewRetr.GetRowCellValue(GridViewRetr.RowCount - 1, "EMPID")?.ToString();

                if (!string.IsNullOrEmpty(sEmpid))
                {
                    GridViewRetr.AddNewRow();
                    GridViewRetr.UpdateCurrentRow();

                    GridViewRetr.SetFocusedRowCellValue("BASDT", DateFrom.EditValue?.ToString().Substring(0, 10));
                    GridViewRetr.FocusedColumn = GridColEmpid;
                }
            }
            SetMagamBtn();
        }
        #endregion

        #region 수정
        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateFrom.EditValue?.ToString().Substring(0, 10)))
            {
                XtraMessageBox.Show("마감된 자료는 수정이 불가능합니다.");
                return;
            }

            DataTable dtRetr = GridRetr.DataSource as DataTable;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dtRetr);
            DataTable dtMerge = dsSave.Tables[0];

            if (dtMerge == null || dtMerge.Rows.Count == 0)
            {
                XtraMessageBox.Show("편집한 데이터가 없습니다.");
                return;
            }

            if (ChkPayApply(dtMerge.Rows[0]["BASDT"]?.ToString().Substring(0,10)))
            {
                XtraMessageBox.Show("급여 승인이 완료되어 근태수정을 할 수 없습니다.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("편집한 데이터를 모두 저장하시겠습니까?"), "저장확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            if (SaveGDAYF(dtMerge))
            {
                BtnRetr.PerformClick();
                GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColEmpid, dtMerge.Rows[0]["EMPID"]?.ToString());

                XtraMessageBox.Show("저장이 완료되었습니다.");
            }
        }

        private bool SaveGDAYF(DataTable dtModi)
        {
            bool result = false;
            string sBASDT  = string.Empty;
            string sEMPID  = string.Empty;
            string sSEQ = string.Empty;
            string sWKGUB  = string.Empty;
            string sWKINTM = string.Empty;
            string sWKOTTM = string.Empty;
            string sGOINTM = string.Empty;
            string sGOOTTM = string.Empty;
            string sWKBASE = string.Empty;
            string sGWKTM1 = string.Empty;
            string sGWKTM2 = string.Empty;
            string sGWKTM3 = string.Empty;
            string sGWKTM4 = string.Empty;
            string sGWKTM5 = string.Empty;
            string sGWKTM6 = string.Empty;
            string sGWKTM7 = string.Empty;
            string sGWKTM8 = string.Empty;
            string sGWKTM9 = string.Empty;

            double dSEQ = 0;

            double dWKBASE = 0.0;
            double dGWKTM1 = 0.0;
            double dGWKTM2 = 0.0;
            double dGWKTM3 = 0.0;
            double dGWKTM4 = 0.0;
            double dGWKTM5 = 0.0;
            double dGWKTM6 = 0.0;
            double dGWKTM7 = 0.0;
            double dGWKTM8 = 0.0;
            double dGWKTM9 = 0.0;

            string sRK1 = string.Empty;
            string sUSER = FmMainToolBar2.drUser["USRCD"]?.ToString();

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < dtModi.Rows.Count; i++)
                {
                    sBASDT = dtModi.Rows[i]["BASDT"]?.ToString();
                    sEMPID = dtModi.Rows[i]["EMPID"]?.ToString();
                    sSEQ = dtModi.Rows[i]["SEQ"]?.ToString();
                    sWKGUB = dtModi.Rows[i]["WKGUB"]?.ToString();
                    sWKINTM = dtModi.Rows[i]["WKINTM"]?.ToString();
                    sWKOTTM = dtModi.Rows[i]["WKOTTM"]?.ToString();
                    sGOINTM = dtModi.Rows[i]["GOINTM"]?.ToString();
                    sGOOTTM = dtModi.Rows[i]["GOOTTM"]?.ToString();
                    
                    sRK1 = dtModi.Rows[i]["RK1"]?.ToString();

                    sWKBASE = dtModi.Rows[i]["WKBASE"]?.ToString();
                    sGWKTM1 = dtModi.Rows[i]["GWKTM1"]?.ToString();
                    sGWKTM2 = dtModi.Rows[i]["GWKTM2"]?.ToString();
                    sGWKTM3 = dtModi.Rows[i]["GWKTM3"]?.ToString();
                    sGWKTM4 = dtModi.Rows[i]["GWKTM4"]?.ToString();
                    sGWKTM5 = dtModi.Rows[i]["GWKTM5"]?.ToString();
                    sGWKTM6 = dtModi.Rows[i]["GWKTM6"]?.ToString();
                    sGWKTM7 = dtModi.Rows[i]["GWKTM7"]?.ToString();
                    sGWKTM8 = dtModi.Rows[i]["GWKTM8"]?.ToString();
                    sGWKTM9 = dtModi.Rows[i]["GWKTM9"]?.ToString();

                    double.TryParse(sSEQ, out dSEQ);

                    double.TryParse(sWKBASE, out dWKBASE);
                    double.TryParse(sGWKTM1, out dGWKTM1);
                    double.TryParse(sGWKTM2, out dGWKTM2);
                    double.TryParse(sGWKTM3, out dGWKTM3);
                    double.TryParse(sGWKTM4, out dGWKTM4);
                    double.TryParse(sGWKTM5, out dGWKTM5);
                    double.TryParse(sGWKTM6, out dGWKTM6);
                    double.TryParse(sGWKTM7, out dGWKTM7);
                    double.TryParse(sGWKTM8, out dGWKTM8);
                    double.TryParse(sGWKTM9, out dGWKTM9);

                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT* FROM GDAYF WHERE BASDT = '"+ sBASDT + "' AND EMPID = '"+ sEMPID + "' AND SEQ = "+ dSEQ + ")  ");
                    strSql.AppendLine("     BEGIN                                                                  ");
                    strSql.AppendLine("           UPDATE GDAYF                                                   ");
                    strSql.AppendLine("              SET WKGUB = '"+ sWKGUB + "'                                                 ");
                    strSql.AppendLine("                , WKINTM = '"+ sWKINTM + "'                                              ");
                    strSql.AppendLine("                , WKOTTM = '"+ sWKOTTM + "'                                              ");
                    strSql.AppendLine("                , GOINTM = '"+ sGOINTM + "'                                              ");
                    strSql.AppendLine("                , GOOTTM = '"+ sGOOTTM + "'                                              ");
                    strSql.AppendLine("                , WKBASE = "+ dWKBASE + "                                               ");
                    strSql.AppendLine("                , GWKTM1 = "+ dGWKTM1 + "                                               ");
                    strSql.AppendLine("                , GWKTM2 = "+ dGWKTM2 + "                                               ");
                    strSql.AppendLine("                , GWKTM3 = "+ dGWKTM3 + "                                               ");
                    strSql.AppendLine("                , GWKTM4 = "+ dGWKTM4 + "                                               ");
                    strSql.AppendLine("                , GWKTM5 = "+ dGWKTM5 + "                                               ");
                    strSql.AppendLine("                , GWKTM6 = "+ dGWKTM6 + "                                               ");
                    strSql.AppendLine("                , GWKTM7 = "+ dGWKTM7 + "                                               ");
                    strSql.AppendLine("                , GWKTM8 = "+ dGWKTM8 + "                                               ");
                    strSql.AppendLine("                , GWKTM9 = "+ dGWKTM9 + "                                               ");
                    strSql.AppendLine("                , RK1 = '"+ sRK1 + "'                                                 ");
                    strSql.AppendLine("                , MDATE = CONVERT(VARCHAR(20), GETDATE(), 20)              ");
                    strSql.AppendLine("                , MUSER = '"+ sUSER + "'                                                 ");
                    strSql.AppendLine("            WHERE BASDT = '"+ sBASDT + "'                                               ");
                    strSql.AppendLine("              AND EMPID = '"+ sEMPID + "'                                                 ");
                    strSql.AppendLine("              AND SEQ = "+ dSEQ + "                                                    ");
                    strSql.AppendLine("         END                                                                ");
                    strSql.AppendLine(" ELSE                                                                       ");
                    strSql.AppendLine("     BEGIN                                                                  ");
                    strSql.AppendLine("         DECLARE @SEQ NUMERIC = 1;                                          ");
                    strSql.AppendLine("          SELECT @SEQ = ISNULL(MAX(SEQ), 0) + 1                  ");
                    strSql.AppendLine("            FROM GDAYF                                                     ");
                    strSql.AppendLine("           WHERE BASDT = '"+ sBASDT + "'                                               ");
                    strSql.AppendLine("             AND EMPID = '"+ sEMPID + "'                                                 ");
                    strSql.AppendLine("                                                                            ");
                    strSql.AppendLine("             INSERT INTO GDAYF(BASDT                                        ");
                    strSql.AppendLine("                             , EMPID                                        ");
                    strSql.AppendLine("                             , SEQ                                          ");
                    strSql.AppendLine("                             , WKGUB                                        ");
                    strSql.AppendLine("                             , WKINTM                                       ");
                    strSql.AppendLine("                             , WKOTTM                                       ");
                    strSql.AppendLine("                             , GOINTM                                       ");
                    strSql.AppendLine("                             , GOOTTM                                       ");
                    strSql.AppendLine("                             , WKBASE                                       ");
                    strSql.AppendLine("                             , GWKTM1                                       ");
                    strSql.AppendLine("                             , GWKTM2                                       ");
                    strSql.AppendLine("                             , GWKTM3                                       ");
                    strSql.AppendLine("                             , GWKTM4                                       ");
                    strSql.AppendLine("                             , GWKTM5                                       ");
                    strSql.AppendLine("                             , GWKTM6                                       ");
                    strSql.AppendLine("                             , GWKTM7                                       ");
                    strSql.AppendLine("                             , GWKTM8                                       ");
                    strSql.AppendLine("                             , GWKTM9                                       ");
                    strSql.AppendLine("                             , RK1                                          ");
                    strSql.AppendLine("                             , CUSER)                                       ");
                    strSql.AppendLine("                       VALUES( '"+ sBASDT + "'                                       ");
                    strSql.AppendLine("                             , '"+ sEMPID + "'                                        ");
                    strSql.AppendLine("                             , @SEQ                                          ");
                    strSql.AppendLine("                             , '"+ sWKGUB + "'                                       ");
                    strSql.AppendLine("                             , '"+ sWKINTM + "'                                       ");
                    strSql.AppendLine("                             , '"+ sWKOTTM + "'                                       ");
                    strSql.AppendLine("                             , '"+ sGOINTM + "'                                       ");
                    strSql.AppendLine("                             , '"+ sGOOTTM + "'                                       ");
                    strSql.AppendLine("                             , "+ dWKBASE + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM1 + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM2 + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM3 + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM4 + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM5 + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM6 + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM7 + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM8 + "                                       ");
                    strSql.AppendLine("                             , "+ dGWKTM9 + "                                       ");
                    strSql.AppendLine("                             , '"+ sRK1 + "'                                          ");
                    strSql.AppendLine("                             , '"+ sUSER + "')                                       ");
                    strSql.AppendLine("         END                                                                ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                result = true;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);

                result = false;
            }

            return result;
        }
        #endregion

        #region 삭제
        private void BtnDel_Click(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateFrom.EditValue?.ToString().Substring(0,10)))
            {
                XtraMessageBox.Show("마감된 자료는 삭제가 불가능합니다.");
                return;
            }

            int[] selRows = GridViewRetr.GetSelectedRows();

            if (selRows.Length == 0)
            {
                XtraMessageBox.Show("삭제할 항목을 선택해주세요.");
                return;
            }

            if (XtraMessageBox.Show(string.Format("정말로 선택한 데이터를 삭제하시겠습니까?"), "삭제확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            DataTable dtRetr = GridRetr.DataSource as DataTable;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dtRetr);
            DataTable dtMerge = dsSave.Tables[0];

            if (dtMerge != null && dtMerge.Rows.Count != 0)
            {
                if (SaveGDAYF(dtMerge))
                {

                }
                else
                {
                    XtraMessageBox.Show("삭제 전 저장 실패");
                    return;
                }
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                for (int i = 0; i < selRows.Length; i++)
                {
                    string sBASDT = GridViewRetr.GetRowCellValue(selRows[i], "BASDT")?.ToString();
                    string sEMPID = GridViewRetr.GetRowCellValue(selRows[i], "EMPID")?.ToString();
                    string sSEQ = GridViewRetr.GetRowCellValue(selRows[i], "SEQ")?.ToString();

                    strSql.Clear();
                    strSql.AppendLine("DELETE FROM GDAYF");
                    strSql.AppendLine(" WHERE BASDT = '" + sBASDT + "'");
                    strSql.AppendLine("   AND EMPID = '" + sEMPID + "'");
                    strSql.AppendLine("   AND SEQ = " + sSEQ);

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

        #region 근태생성
        private void BtnCrate_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder strSql = new StringBuilder();

                string sDate = DateFrom.EditValue?.ToString().Substring(0, 10);

                if (ChkMagamYN(sDate))
                {
                    XtraMessageBox.Show("마감된 자료는 근태생성이 불가능합니다.");
                    return;
                }

                if (ChkPayApply(sDate))
                {
                    XtraMessageBox.Show("급여 승인이 완료되어 근태생성을 할 수 없습니다.");
                    return;
                }

                strSql.Clear();
                strSql.AppendLine(" WITH TEMP1 AS(                                                 ");
                strSql.AppendLine("     --출근                                                     ");
                strSql.AppendLine("     SELECT CONVERT(DATE, E_DATE) AS BASDT                      ");
                strSql.AppendLine("          , E_ID                                                ");
                strSql.AppendLine("          , E_IDNO                                              ");
                strSql.AppendLine("          , STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS ETIME ");
                strSql.AppendLine("          , CONVERT(VARCHAR, CONVERT(DATE, E_DATE)) + ' ' + STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS EDATETIME");
                strSql.AppendLine("          , E_MODE                        ");
                strSql.AppendLine("       FROM TENTER                        ");
                strSql.AppendLine("      WHERE E_DATE = '" + sDate.Replace("-", "") + "'           ");
                strSql.AppendLine("        AND E_MODE = '1'                  ");
                strSql.AppendLine(" ), TEMP2 AS(                             ");
                strSql.AppendLine("     --퇴근                               ");
                strSql.AppendLine("     SELECT CONVERT(DATE, E_DATE) AS BASDT");
                strSql.AppendLine("          , E_ID                          ");
                strSql.AppendLine("          , E_IDNO                        ");
                strSql.AppendLine("          , STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS ETIME");
                strSql.AppendLine("          , CONVERT(VARCHAR, CONVERT(DATE, E_DATE)) +' ' + STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS EDATETIME");
                strSql.AppendLine("           , E_MODE              ");
                strSql.AppendLine("       FROM TENTER               ");
                strSql.AppendLine("      WHERE E_DATE = '" + sDate.Replace("-", "") + "'  ");
                strSql.AppendLine("        AND E_MODE = '2'         ");
                strSql.AppendLine(" ), TEMP3 AS(                             ");
                strSql.AppendLine("     --외출                               ");
                strSql.AppendLine("     SELECT CONVERT(DATE, E_DATE) AS BASDT");
                strSql.AppendLine("          , E_ID                          ");
                strSql.AppendLine("          , E_IDNO                        ");
                strSql.AppendLine("          , STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS ETIME");
                strSql.AppendLine("          , CONVERT(VARCHAR, CONVERT(DATE, E_DATE)) +' ' + STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS EDATETIME");
                strSql.AppendLine("           , E_MODE              ");
                strSql.AppendLine("       FROM TENTER               ");
                strSql.AppendLine("      WHERE E_DATE = '" + sDate.Replace("-", "") + "'  ");
                strSql.AppendLine("        AND E_MODE = '3'         ");
                strSql.AppendLine(" ), TEMP4 AS(                             ");
                strSql.AppendLine("     --복귀                               ");
                strSql.AppendLine("     SELECT CONVERT(DATE, E_DATE) AS BASDT");
                strSql.AppendLine("          , E_ID                          ");
                strSql.AppendLine("          , E_IDNO                        ");
                strSql.AppendLine("          , STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS ETIME");
                strSql.AppendLine("          , CONVERT(VARCHAR, CONVERT(DATE, E_DATE)) +' ' + STUFF(STUFF(E_TIME, 3, 0, ':'), 6, 0, ':') AS EDATETIME");
                strSql.AppendLine("           , E_MODE              ");
                strSql.AppendLine("       FROM TENTER               ");
                strSql.AppendLine("      WHERE E_DATE = '" + sDate.Replace("-", "") + "'  ");
                strSql.AppendLine("        AND E_MODE = '4'         ");
                strSql.AppendLine(" )                               ");
                strSql.AppendLine("                                 ");
                strSql.AppendLine("  SELECT CONVERT(DATE, '" + sDate.Replace("-", "") + "') AS BASDT");
                strSql.AppendLine("       , A1.EMP_ID          ");
                strSql.AppendLine("       , ROW_NUMBER() OVER(PARTITION BY A1.EMP_ID ORDER BY B1.ETIME,B2.ETIME,B3.ETIME,B4.ETIME) AS SEQ          ");
                strSql.AppendLine("       , A1.EMP_NM          ");
                strSql.AppendLine("       , A1.GENDER_CD ");
                strSql.AppendLine("       , A1.JMUNYN");
                strSql.AppendLine("       , B1.ETIME AS SCMTM1 ");
                strSql.AppendLine("       , B2.ETIME AS SCMTM2 ");
                strSql.AppendLine("       , B3.ETIME AS SCMTM3 ");
                strSql.AppendLine("       , B4.ETIME AS SCMTM4 ");
                strSql.AppendLine("    FROM HR_EMP_BASIS A1       ");
                strSql.AppendLine("    LEFT JOIN TEMP1 B1         ");
                strSql.AppendLine("      ON A1.EMP_ID = B1.E_IDNO ");
                strSql.AppendLine("    LEFT JOIN TEMP2 B2         ");
                strSql.AppendLine("      ON A1.EMP_ID = B2.E_IDNO ");
                strSql.AppendLine("    LEFT JOIN TEMP3 B3         ");
                strSql.AppendLine("      ON A1.EMP_ID = B3.E_IDNO ");
                strSql.AppendLine("    LEFT JOIN TEMP4 B4         ");
                strSql.AppendLine("      ON A1.EMP_ID = B4.E_IDNO ");
                strSql.AppendLine("   WHERE A1.EMPL_GB = 'Y'      ");
                strSql.AppendLine("     AND A1.GDAYYN = 'Y'       ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if(dt != null && dt.Rows.Count > 0)
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    for(int i=0;i<dt.Rows.Count; i++)
                    {
                        string sBASDT  = dt.Rows[i]["BASDT"]?.ToString().Substring(0,10);
                        string sEMPID  = dt.Rows[i]["EMP_ID"]?.ToString();
                        string sSEQ = dt.Rows[i]["SEQ"]?.ToString();
                        string sSCMTM1 = dt.Rows[i]["SCMTM1"]?.ToString();
                        string sSCMTM2 = dt.Rows[i]["SCMTM2"]?.ToString();
                        string sSCMTM3 = dt.Rows[i]["SCMTM3"]?.ToString();
                        string sSCMTM4 = dt.Rows[i]["SCMTM4"]?.ToString();
                        string sGENDER = dt.Rows[i]["GENDER_CD"]?.ToString();
                        string sJMUNYN = dt.Rows[i]["JMUNYN"]?.ToString();
                        string sUser = FmMainToolBar2.drUser["USRCD"]?.ToString();

                        string sWKGUB = string.Empty;

                        double dAlWKBS = 0.0;//총근무시간
                        double dWKBASE = 0.0;//근무시간
                        double dGWKTM2 = 0.0;//기본잔업 남직원:1.5 여직원:0.0 2022-08-24 요청
                        double dGWKTM3 = 0.0;//추가잔업 
                        double dGWKTM5 = 0.0;//지각시간
                        //double dGWKTM6 = 0.0;//조퇴시간
                        double dGWKTM7 = 0.0;//외출시간

                        DateTime chulTime = new DateTime();//출근시간
                        DateTime chulTime2 = new DateTime();//퇴근시간
                        DateTime janTime = DateTime.Parse(sBASDT + " " + _NTBAB);//잔업시작시간

                        if (!string.IsNullOrEmpty(sGENDER) && sGENDER.Equals("1"))//남직원. 일반
                        {
                            chulTime = DateTime.Parse(sBASDT + " " + _ILBAN1);
                            chulTime2 = DateTime.Parse(sBASDT + " " + _ILBAN2);


                            if(chulTime.DayOfWeek != DayOfWeek.Saturday && chulTime.DayOfWeek != DayOfWeek.Sunday)
                            {
                                //지문여부 N 일때 기본 출근시간
                                if (string.IsNullOrEmpty(sSCMTM1) && !string.IsNullOrEmpty(sJMUNYN) && sJMUNYN.Equals("N"))
                                {
                                    sSCMTM1 = _ILBAN1;
                                }

                                //지문여부 N 일때 기본 퇴근시간
                                if (string.IsNullOrEmpty(sSCMTM2) && !string.IsNullOrEmpty(sJMUNYN) && sJMUNYN.Equals("N"))
                                {
                                    sSCMTM2 = _ILBAN2;
                                }
                            }
                        }
                        else if (!string.IsNullOrEmpty(sGENDER) && sGENDER.Equals("2"))//여직원
                        {
                            chulTime = DateTime.Parse(sBASDT + " " + _YUKIC1);
                            chulTime2 = DateTime.Parse(sBASDT + " " + _YUKIC2);

                            if (chulTime.DayOfWeek != DayOfWeek.Saturday && chulTime.DayOfWeek != DayOfWeek.Sunday)
                            {
                                //지문여부 N 일때 기본 출근시간
                                if (string.IsNullOrEmpty(sSCMTM1) && !string.IsNullOrEmpty(sJMUNYN) && sJMUNYN.Equals("N"))
                                {
                                    sSCMTM1 = _YUKIC1;
                                }

                                //지문여부 N 일때 기본 퇴근시간
                                if (string.IsNullOrEmpty(sSCMTM2) && !string.IsNullOrEmpty(sJMUNYN) && sJMUNYN.Equals("N"))
                                {
                                    sSCMTM2 = _YUKIC2;
                                }
                            }
                        }

                        //출근시간 비교 출근시간 = 7:00
                        DateTime.TryParse(sBASDT + " " + sSCMTM1, out DateTime dtTime);

                        //퇴근시간 비교 퇴근시간 = 17:30
                        DateTime.TryParse(sBASDT + " " + sSCMTM2, out DateTime dtTime2);

                        if (dtTime.Hour > _YAGAN)
                        {
                            //야간근무인 경우 퇴근 다음날
                            dtTime2 = dtTime2.AddDays(1);
                        }

                        if (!string.IsNullOrEmpty(sSCMTM1))
                        {

                            if (dtTime.Hour < _YAGAN && chulTime < dtTime) //야근조가 아니고(야근조: 18시 이후 출근), 지각
                            {
                                TimeSpan ts = dtTime - chulTime;

                                //지각
                                dGWKTM5 = (double)((ts.Hours * 60) + ts.Minutes) / 60;

                                //#0001
                                double dTemp = dGWKTM5 - Math.Truncate(dGWKTM5);
                                if(dTemp > 0 && dTemp < 0.5)
                                {
                                    dGWKTM5 = dGWKTM5 - dTemp + 0.5;
                                }
                                else if(dTemp > 0 && dTemp > 0.5)
                                {
                                    dGWKTM5 = dGWKTM5 - dTemp + 1.0;
                                }
                            }
                            else if (dtTime.Hour > _YAGAN)
                            {
                            }
                        }

                        if (!string.IsNullOrEmpty(sSCMTM2))
                        {

                            if (!string.IsNullOrEmpty(sSCMTM1))
                            {
                                if (!string.IsNullOrEmpty(sGENDER) && sGENDER.Equals("1"))//남직원. 일반
                                {
                                    dGWKTM2 = 1.5;
                                }
                                else if (!string.IsNullOrEmpty(sGENDER) && sGENDER.Equals("2"))//여직원
                                {
                                    dGWKTM2 = 0.0;
                                }

                                //총근무시간
                                TimeSpan ts = dtTime2 - dtTime;
                                dAlWKBS = Math.Round((double)(((ts.Hours * 60) + ts.Minutes) / 60.0), 1, MidpointRounding.AwayFromZero);
                                dAlWKBS = Math.Round(dAlWKBS * 2, MidpointRounding.AwayFromZero) / 2;

                                if (dAlWKBS >= 8)
                                {
                                    dWKBASE = 8;
                                }
                                else
                                {
                                    dWKBASE = dAlWKBS;
                                }

                                //추가잔업 18시 이후 근무시간. 야간근무제외
                                if (dtTime2 > janTime && dtTime.Hour < _YAGAN)
                                {
                                    TimeSpan tsJan = dtTime2 - janTime;

                                    //dGWKTM2 = 1.5;
                                    dGWKTM3 = Math.Round((double)(((tsJan.Hours * 60) + tsJan.Minutes) / 60.0), 1, MidpointRounding.AwayFromZero);
                                    dGWKTM3 = Math.Round(dGWKTM3 * 2, MidpointRounding.AwayFromZero) / 2;

                                    if (dGWKTM3 > 0)
                                    {
                                    }
                                }
                                //else if(dtTime2 >= chulTime2 && dtTime.Hour < _YAGAN) //추가잔업x 야간근무x
                                //{
                                //    dGWKTM3 = 0;
                                //    dGWKTM2 = dAlWKBS - 9;

                                //    if(dGWKTM2 > 1.5)
                                //    {
                                //        dGWKTM2 = 1.5;
                                //    }
                                //    else if (dGWKTM2 < 0)
                                //    {
                                //        dGWKTM2 = 0;
                                //    }
                                //    else
                                //    {
                                //        dGWKTM2 = Math.Round(dGWKTM2 * 2, MidpointRounding.AwayFromZero) / 2;
                                //    }
                                //}

                                //조퇴(야간제외) //수기입력으로 한다함 2022-08-24
                                //if(dtTime2 < chulTime2 && dtTime.Hour < _YAGAN) //조퇴
                                //{
                                //    TimeSpan tsJote = chulTime2 - dtTime2;

                                //    dGWKTM6 = Math.Round(((double)((tsJote.Hours * 60) + tsJote.Minutes) / 60), 1, MidpointRounding.AwayFromZero);
                                //    dGWKTM6 = Math.Round(dGWKTM6 * 2, MidpointRounding.AwayFromZero) / 2;

                                //    if(dGWKTM6 > 0)
                                //    {
                                //    }
                                //}
                            }
                        }

                        if (DateTime.TryParse(sSCMTM3, out DateTime time3) && DateTime.TryParse(sSCMTM3, out DateTime time4))
                        {
                            //외출시간
                            TimeSpan tsout = time4 - time3;

                            dGWKTM7 = (double)((tsout.Hours * 60) + tsout.Minutes) / 60;

                            //#0001
                            double dTemp = dGWKTM5 - Math.Truncate(dGWKTM5);
                            if (dTemp > 0 && dTemp < 0.5)
                            {
                                dGWKTM5 = dGWKTM5 - dTemp + 0.5;
                            }
                            else if (dTemp > 0 && dTemp > 0.5)
                            {
                                dGWKTM5 = dGWKTM5 - dTemp + 1.0;
                            }
                        }

                        strSql.Clear();
                        strSql.AppendLine(" IF EXISTS(SELECT * FROM GDAYF WHERE BASDT = '" + sBASDT + "' AND EMPID = '" + sEMPID + "' AND SEQ = "+ sSEQ + ") ");
                        strSql.AppendLine("      BEGIN                                                    ");
                        strSql.AppendLine("            UPDATE GDAYF                                       ");
                        strSql.AppendLine("               SET WKINTM = '" + sSCMTM1 + "'");
                        strSql.AppendLine("                 , WKOTTM = '" + sSCMTM2 + "'");
                        strSql.AppendLine("                 , GOINTM = '"+ sSCMTM3 + "'");
                        strSql.AppendLine("                 , GOOTTM = '"+sSCMTM4+"'");
                        strSql.AppendLine("                 , WKBASE = "+ dWKBASE + " ");
                        strSql.AppendLine("                 , GWKTM2 = "+ dGWKTM2);
                        strSql.AppendLine("                 , GWKTM3 = "+ dGWKTM3 + " ");
                        strSql.AppendLine("                 , GWKTM5 = " + dGWKTM5 + "");
                        strSql.AppendLine("                 , GWKTM7 = " + dGWKTM7);
                        strSql.AppendLine("                 , MUSER = '" + sUser + "'                                  ");
	                    strSql.AppendLine("                 , MDATE = CONVERT(VARCHAR(20), GETDATE(), 20) ");
                        strSql.AppendLine("             WHERE BASDT = '" + sBASDT + "'                                  ");
                        strSql.AppendLine("               AND EMPID = '" + sEMPID + "'                                  ");
                        strSql.AppendLine("               AND SEQ = " + sSEQ);
                        strSql.AppendLine("        END                                                    ");
                        strSql.AppendLine(" ELSE                                                          ");
                        strSql.AppendLine("      BEGIN                                                    ");
                        strSql.AppendLine("            INSERT INTO GDAYF(BASDT   ");
                        strSql.AppendLine("                            , EMPID   ");
                        strSql.AppendLine("                            , SEQ   ");
                        strSql.AppendLine("                            , WKINTM ");
                        strSql.AppendLine("                            , WKOTTM ");
                        strSql.AppendLine("                            , GOINTM");
                        strSql.AppendLine("                            , GOOTTM");
                        strSql.AppendLine("                            , WKBASE ");
                        strSql.AppendLine("                            , GWKTM2");
                        strSql.AppendLine("                            , GWKTM3 ");
                        strSql.AppendLine("                            , GWKTM5");
                        strSql.AppendLine("                            , GWKTM7");
                        strSql.AppendLine("                            , CUSER ) ");
                        strSql.AppendLine("                      VALUES( '" + sBASDT + "'      ");
                        strSql.AppendLine("                            , '" + sEMPID + "'      ");
                        strSql.AppendLine("                            , " + sSEQ);
                        strSql.AppendLine("                            , '" + sSCMTM1 + "'      ");
                        strSql.AppendLine("                            , '" + sSCMTM2 + "'      ");
                        strSql.AppendLine("                            , '"+ sSCMTM3 + "'");
                        strSql.AppendLine("                            , '"+sSCMTM4+"'");
                        strSql.AppendLine("                            , " + dWKBASE + "  ");
                        strSql.AppendLine("                            , " + dGWKTM2);
                        strSql.AppendLine("                            , " + dGWKTM3 + "  ");
                        strSql.AppendLine("                            , " + dGWKTM5);
                        strSql.AppendLine("                            , " + dGWKTM7);
                        strSql.AppendLine("                            , '" + sUser + "' )    ");
                        strSql.AppendLine("        END                                        ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    XtraMessageBox.Show("지문데이터 생성을 완료 했습니다.");
                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;

                    BtnRetr.PerformClick();
                }
            }                                                                                 
            catch(Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 마감
        //마감확인
        private bool ChkMagamYN(string sDate)
        {
            bool result = false;
            try
            {

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT COUNT(*) AS CNT     ");
                strSql.AppendLine("   FROM GDAYF               ");
                strSql.AppendLine("  WHERE BASDT = '" + sDate + "'");
                strSql.AppendLine("    AND MCHKYN = 'Y'        ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt != null)
                {
                    double iCnt = 0;
                    double.TryParse(dt.Rows[0]["CNT"]?.ToString(), out iCnt);

                    if (iCnt > 0) //마감된 자료가 있으면 마감됨 true
                    {
                        result = true;
                    }
                    else
                    {
                        result = false;
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString(), "조회 오류");
            }

            return result;
        }

        private void BtnMagam_Click(object sender, EventArgs e)
        {
            DataTable dtRetr = GridRetr.DataSource as DataTable;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dtRetr);
            DataTable dtMerge = dsSave.Tables[0];

            if (dtMerge != null && dtMerge.Rows.Count != 0)
            {
                if (SaveGDAYF(dtMerge))
                {

                }
                else
                {
                    XtraMessageBox.Show("마감 전 저장 실패");
                    return;
                }
            }

            if (BtnMagam.Text.Equals("마감"))
            {
                Magam();
            }
            else
            {
                MagamCancle();
            }
        }
        
        private void Magam()
        {
            string sDate = DateFrom.EditValue?.ToString().Substring(0, 10);

            if (XtraMessageBox.Show(string.Format(sDate + "일자의\r\n데이터를 마감하시겠습니까?"), "마감확인", MessageBoxButtons.YesNo) == DialogResult.No)
                return;

            try
            {
                StringBuilder strSql = new StringBuilder();

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine(" UPDATE GDAYF        ");
                strSql.AppendLine("    SET MCHKYN = 'Y' ");
                strSql.AppendLine("  WHERE BASDT = '"+sDate+"'   ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                XtraMessageBox.Show("마감을 완료했습니다.");
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                BtnRetr.PerformClick();
            }
            catch(Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                XtraMessageBox.Show(ex.Message);
            }
        }

        private void MagamCancle()
        {
            try
            {
                string sDate = DateFrom.EditValue?.ToString().Substring(0, 10);

                if (ChkPayApply(sDate))
                {
                    XtraMessageBox.Show("급여 승인이 완료된 근태 데이터는 마감취소 할 수 없습니다.");
                    return;
                }

                StringBuilder strSql = new StringBuilder();

                if (XtraMessageBox.Show(string.Format(sDate + "일자의\r\n데이터를 마감취소하시겠습니까?"), "마감취소확인", MessageBoxButtons.YesNo) == DialogResult.No)
                    return;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                strSql.Clear();
                strSql.AppendLine(" UPDATE GDAYF        ");
                strSql.AppendLine("    SET MCHKYN = 'N' ");
                strSql.AppendLine("  WHERE BASDT = '" + sDate + "'   ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                XtraMessageBox.Show("마감취소를 완료했습니다.");
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                BtnRetr.PerformClick();
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                XtraMessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region 급여 승인 체크
        private bool ChkPayApply(string sDate)
        {
            bool result = false;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT COUNT(*) AS CNT            ");
            strSql.AppendLine("   FROM PMONF_1           ");
            strSql.AppendLine("  WHERE BASYM = '" + sDate.Substring(0, 7) + "' ");
            strSql.AppendLine("    AND PAPLY = 'Y'");

            DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dtChk != null && dtChk.Rows.Count > 0)
            {
                if (int.Parse(dtChk.Rows[0]["CNT"]?.ToString()) > 0)//급여 마감처리된 경우
                {
                    result = true;
                }
            }

            return result;
        }
        #endregion

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GE001F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F1)
                BtnAdd.PerformClick();
            else if (e.KeyCode == Keys.F3)
                BtnSave.PerformClick();
            else if (e.KeyCode == Keys.F4)
                BtnDel.PerformClick();
            else if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
        }

        private void TxtNmk_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewRetr_CustomDrawRowIndicator(object sender, DevExpress.XtraGrid.Views.Grid.RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        #region 출퇴근관련 입력 이벤트
        //출퇴근구분 선택 후 출퇴근시간 입력하면 자동으로 지각, 조퇴
        private void RepoTxtTktm1_Leave(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateFrom.EditValue?.ToString().Substring(0, 10)))
            {
                return;
            }

            TextEdit textEdit = (TextEdit)sender;

            if (string.IsNullOrEmpty(textEdit.EditValue?.ToString()))
            {
                GridViewRetr.SetFocusedRowCellValue(GridColWkbas, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm2, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm3, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm4, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm5, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm6, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm7, 0.0);
            }
            else
            {
            }

            if(DateTime.TryParse(textEdit.EditValue?.ToString(), out DateTime dt))
                GridViewRetr.SetFocusedRowCellValue(GridColTktm1, dt.ToString("HH:mm"));

            setWkbas();
        }

        private void RepoTxtTktm2_Leave(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateFrom.EditValue?.ToString().Substring(0, 10)))
            {
                return;
            }

            TextEdit textEdit = (TextEdit)sender;

            if (string.IsNullOrEmpty(textEdit.EditValue?.ToString()))
            {
                GridViewRetr.SetFocusedRowCellValue(GridColWkbas, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm2, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm3, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm4, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm5, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm6, 0.0);
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm7, 0.0);
            }
            else
            {
            }

            if(DateTime.TryParse(textEdit.EditValue?.ToString(), out DateTime dt))
                GridViewRetr.SetFocusedRowCellValue(GridColTktm2, dt.ToString("HH:mm"));

            setWkbas();
        }

        private void setWkbas()
        {
            string sBASDT = GridViewRetr.GetFocusedRowCellValue("BASDT")?.ToString();
            string sWKINTM = GridViewRetr.GetFocusedRowCellValue("WKINTM")?.ToString();
            string sWKOTTM = GridViewRetr.GetFocusedRowCellValue("WKOTTM")?.ToString();
            string sWkgub = GridViewRetr.GetFocusedRowCellValue("WKGUB")?.ToString();
            string sGENDER = GridViewRetr.GetFocusedRowCellValue("GENDER_CD")?.ToString();

            double dAlWKBS = 0.0;//총근무시간
            double dWKBASE = 0.0;//근무시간
            double dGWKTM2 = 0.0;//기본잔업 남직원:1.5 여직원:0.0
            double dGWKTM3 = 0.0;//추가잔업 
            double dGWKTM4 = 0.0;//당직
            double dGWKTM5 = 0.0;//지각시간
            //double dGWKTM6 = 0.0;//조퇴시간

            DateTime chulTime = new DateTime();//출근시간
            DateTime chulTime2 = new DateTime();//퇴근시간
            DateTime janTime = DateTime.Parse(sBASDT + " " + _NTBAB);//잔업시작시간

            if (!string.IsNullOrEmpty(sGENDER) && sGENDER.Equals("1"))//남직원. 일반
            {
                chulTime = DateTime.Parse(sBASDT + " " + _ILBAN1);
                chulTime2 = DateTime.Parse(sBASDT + " " + _ILBAN2);

                dGWKTM2 = 1.5;
            }
            else if (!string.IsNullOrEmpty(sGENDER) && sGENDER.Equals("2"))//여직원
            {
                chulTime = DateTime.Parse(sBASDT + " " + _YUKIC1);
                chulTime2 = DateTime.Parse(sBASDT + " " + _YUKIC2);

                dGWKTM2 = 0.0;
            }

            //출근시간 비교 출근시간 = 7:00
            DateTime.TryParse(sBASDT + " " + sWKINTM, out DateTime dtTime);

            //퇴근시간 비교 퇴근시간 = 17:30
            DateTime.TryParse(sBASDT + " " + sWKOTTM, out DateTime dtTime2);

            if (dtTime.Hour > _YAGAN)
            {
                //야간근무인 경우 퇴근 다음날
                dtTime2 = dtTime2.AddDays(1);
            }

            if (!string.IsNullOrEmpty(sWKINTM))
            {

                if (dtTime.Hour < _YAGAN && chulTime < dtTime) //야근조가 아니고(야근조: 18시 이후 출근), 지각
                {
                    TimeSpan ts = dtTime - chulTime;

                    //지각
                    dGWKTM5 = (double)((ts.Hours * 60) + ts.Minutes) / 60;

                    //#0001
                    double dTemp = dGWKTM5 - Math.Truncate(dGWKTM5);
                    if (dTemp > 0 && dTemp < 0.5)
                    {
                        dGWKTM5 = dGWKTM5 - dTemp + 0.5;
                    }
                    else if (dTemp > 0 && dTemp > 0.5)
                    {
                        dGWKTM5 = dGWKTM5 - dTemp + 1.0;
                    }
                }
                else if (dtTime.Hour > _YAGAN)
                {
                }
            }

            if (!string.IsNullOrEmpty(sWKOTTM))
            {

                if (!string.IsNullOrEmpty(sWKINTM))
                {
                    //총근무시간
                    TimeSpan ts = dtTime2 - dtTime;
                    dAlWKBS = Math.Round((double)(((ts.Hours * 60) + ts.Minutes) / 60.0), 1, MidpointRounding.AwayFromZero);
                    dAlWKBS = Math.Round(dAlWKBS * 2, MidpointRounding.AwayFromZero) / 2;

                    if (dAlWKBS >= 8)
                    {
                        dWKBASE = 8;
                    }
                    else
                    {
                        dWKBASE = dAlWKBS;
                    }

                    //추가잔업 18시 이후 근무시간. 야간근무제외
                    if (dtTime2 > janTime && dtTime.Hour < _YAGAN)
                    {
                        TimeSpan tsJan = dtTime2 - janTime;

                        //dGWKTM2 = 1.5;
                        dGWKTM3 = Math.Round((double)(((tsJan.Hours * 60) + tsJan.Minutes) / 60.0), 1, MidpointRounding.AwayFromZero);
                        dGWKTM3 = Math.Round(dGWKTM3 * 2, MidpointRounding.AwayFromZero) / 2;

                        if (dGWKTM3 > 0)
                        {
                            if (!string.IsNullOrEmpty(sWkgub) && sWkgub.Equals("DJ")) //당직
                            {
                                dGWKTM4 = dGWKTM3;
                                dGWKTM3 = 0;
                            }
                            else
                            {
                            } 
                        }
                    }
                    //else if (dtTime2 >= chulTime2 && dtTime.Hour < _YAGAN) //추가잔업x 야간근무x
                    //{
                    //    dGWKTM3 = 0;
                    //    dGWKTM2 = dAlWKBS - 9;

                    //    if (dGWKTM2 > 1.5)
                    //    {
                    //        dGWKTM2 = 1.5;
                    //    }
                    //    else if (dGWKTM2 < 0)
                    //    {
                    //        dGWKTM2 = 0;
                    //    }
                    //    else
                    //    {
                    //        dGWKTM2 = Math.Round(dGWKTM2 * 2, MidpointRounding.AwayFromZero) / 2;
                    //    }
                    //}

                    //조퇴(야간제외) //수기입력한다함 2022-08-24
                    //if (dtTime2 < chulTime2 && dtTime.Hour < _YAGAN) //조퇴
                    //{
                    //    TimeSpan tsJote = chulTime2 - dtTime2;

                    //    dGWKTM6 = Math.Round(((double)((tsJote.Hours * 60) + tsJote.Minutes) / 60), 1, MidpointRounding.AwayFromZero);
                    //    dGWKTM6 = Math.Round(dGWKTM6 * 2, MidpointRounding.AwayFromZero) / 2;

                    //    if (dGWKTM6 > 0)
                    //    {
                    //    }
                    //}
                }
            }

            GridViewRetr.SetFocusedRowCellValue("WKBASE", dWKBASE);//근무시간
            GridViewRetr.SetFocusedRowCellValue("GWKTM2", dGWKTM2);//기본잔업 1.5
            GridViewRetr.SetFocusedRowCellValue("GWKTM3", dGWKTM3);//추가잔업 
            GridViewRetr.SetFocusedRowCellValue("GWKTM4", dGWKTM4);//당직
            GridViewRetr.SetFocusedRowCellValue("GWKTM5", dGWKTM5);//지각시간
            //GridViewRetr.SetFocusedRowCellValue("GWKTM6", dGWKTM6);//조퇴시간
        }
        #endregion

        #region 외출
        private void RepoTxtGotm1_Leave(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateFrom.EditValue?.ToString().Substring(0, 10)))
            {
                return;
            }

            TextEdit textEdit = (TextEdit)sender;

            if (string.IsNullOrEmpty(textEdit.EditValue?.ToString()))
            {
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm7, 0.0);
                return;
            }

            DateTime dt = DateTime.Parse(textEdit.EditValue?.ToString());
            GridViewRetr.SetFocusedRowCellValue(GridColGotm1, dt.ToString("HH:mm"));

            setGwktm7();
        }

        private void RepoTxtGotm2_Leave(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateFrom.EditValue?.ToString().Substring(0, 10)))
            {
                return;
            }

            TextEdit textEdit = (TextEdit)sender;

            if (string.IsNullOrEmpty(textEdit.EditValue?.ToString()))
            {
                GridViewRetr.SetFocusedRowCellValue(GridColGwktm7, 0.0);
                return;
            }

            DateTime dt = DateTime.Parse(textEdit.EditValue?.ToString());
            GridViewRetr.SetFocusedRowCellValue(GridColGotm2, dt.ToString("HH:mm"));

            setGwktm7();
        }

        private void setGwktm7()
        {
            string sTktm1 = GridViewRetr.GetFocusedRowCellValue(GridColGotm1)?.ToString();
            string sTktm2 = GridViewRetr.GetFocusedRowCellValue(GridColGotm2)?.ToString();

            if (string.IsNullOrEmpty(sTktm1) || string.IsNullOrEmpty(sTktm2))
                return;

            DateTime dtTktm1 = DateTime.Parse(sTktm1);
            DateTime dtTktm2 = DateTime.Parse(sTktm2);

            TimeSpan resultTime = dtTktm2 - dtTktm1;

            double iGwktm7 = 0;
            iGwktm7 = (double)(((resultTime.Hours * 60) + resultTime.Minutes) / 60.0);

            //#0001
            double dTemp = iGwktm7 - Math.Truncate(iGwktm7);
            if (dTemp > 0 && dTemp < 0.5)
            {
                iGwktm7 = iGwktm7 - dTemp + 0.5;
            }
            else if (dTemp > 0 && dTemp > 0.5)
            {
                iGwktm7 = iGwktm7 - dTemp + 1.0;
            }
        }
        #endregion

        #region 출퇴근구분 연차,병가,반차 선택 (사용x)
        private void RepoGridLkupSrtgb_EditValueChanged(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateFrom.EditValue?.ToString().Substring(0, 10)))
            {
                return;
            }

            LookUpEditBase lookUpEdit = (LookUpEditBase)sender;

            string sVal = lookUpEdit.EditValue?.ToString();

            if (sVal == null)
                return;

            int focRowHandle = GridViewRetr.FocusedRowHandle;

            if (sVal.Equals("YG") || sVal.Equals("BG"))//연차개인, 병가 : 추가 연차(5~4)에서 차감
            {
                GridViewRetr.SetRowCellValue(focRowHandle, GridColGwktm1, 1.0);
            }
            else if (sVal.Equals("YA"))//오전반차 : 추가 연차(5~4)에서 차감
            {
                GridViewRetr.SetRowCellValue(focRowHandle, GridColGwktm1, 0.5);
            }
            else
            {
                GridViewRetr.SetRowCellValue(focRowHandle, GridColGwktm1, null);
            }

            GridViewRetr.SetRowCellValue(focRowHandle, GridViewRetr.FocusedColumn, sVal);
        }

        private void RepoGridLkupEndGb_EditValueChanged(object sender, EventArgs e)
        {
            if (ChkMagamYN(DateFrom.EditValue?.ToString().Substring(0, 10)))
            {
                return;
            }

            LookUpEditBase lookUpEdit = (LookUpEditBase)sender;

            string sVal = lookUpEdit.EditValue?.ToString();

            if (sVal == null)
                return;

            int focRowHandle = GridViewRetr.FocusedRowHandle;

            if (sVal.Equals("YP"))//오후반차 : 추가 연차(5~4)에서 차감
            {
                GridViewRetr.SetRowCellValue(focRowHandle, GridColGwktm1, 0.5);
            }
            else
            {
                GridViewRetr.SetRowCellValue(focRowHandle, GridColGwktm1, null);
            }

            GridViewRetr.SetRowCellValue(focRowHandle, GridViewRetr.FocusedColumn, sVal);
        }
        #endregion

        private void BtnMig_Click(object sender, EventArgs e)
        {
            string strConn = "Server=192.168.0.202;Database=daejierp;Uid=daeji;Pwd=dj2019;Charset=UTF8;allow user variables=true;Allow Zero Datetime=True;";//실서버 admin 122
            string sDate = DateFrom.EditValue?.ToString().Substring(0, 10);
            string json = string.Empty;

            DataSet ds = new DataSet();
            DataTable dtMysql = new DataTable();
            StringBuilder strSql = new StringBuilder();
            Dictionary<string, string> dicParams = new Dictionary<string, string>();

            try
            {
                MySqlConnection dbCon = new MySqlConnection(strConn);

                dbCon.Open();

                strSql.Clear();
                strSql.AppendLine(" SELECT E_DATE  ");
                strSql.AppendLine("      , E_TIME  ");
                strSql.AppendLine("      , G_ID    ");
                strSql.AppendLine("      , E_ID    ");
                strSql.AppendLine("      , E_NAME  ");
                strSql.AppendLine("      , E_IDNO  ");
                strSql.AppendLine("      , E_GROUP ");
                strSql.AppendLine("      , E_USER  ");
                strSql.AppendLine("      , E_MODE  ");
                strSql.AppendLine("      , E_TYPE  ");
                strSql.AppendLine("      , E_RESULT");
                strSql.AppendLine("   FROM TENTER  ");
                strSql.AppendLine("  WHERE E_DATE = '" + sDate .Replace("-","")+ "'");

                MySqlDataAdapter adpt = new MySqlDataAdapter(strSql.ToString(), dbCon);

                MySqlTransaction dbTran = adpt.SelectCommand.Transaction;

                adpt.Fill(ds);

                dtMysql = ds.Tables[0];

                if(dtMysql != null)
                {
                    json = ComnEtcFunc.DataTableToJsonObj(dtMysql);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }

            if (!string.IsNullOrEmpty(json))
            {
                dicParams.Clear();
                dicParams.Add("CMD", "GNTAE");
                dicParams.Add("JSON", json);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, "DP_MIG", dicParams);

                if(dt != null)
                {
                    XtraMessageBox.Show(dt.Rows[0]["MSG"]?.ToString());
                }
            }
        }

        private void RepoGridLkupWkgub_EditValueChanged(object sender, EventArgs e)
        {
            LookUpEditBase lookUpEdit = (LookUpEditBase) sender;

            string sVal = lookUpEdit.EditValue?.ToString();

            if (string.IsNullOrEmpty(sVal))
                return;


            if (sVal.Equals("DH")) //대체휴무인 경우 기본 근무시간 8, 기본잔업시간 1.5
            {
                GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "WKBASE", 8);
                GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "GWKTM2", 1.5);
            }
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sEDIT = GridViewRetr.GetFocusedRowCellValue("EDIT")?.ToString();

            if (string.IsNullOrEmpty(sEDIT))
            {
                GridColEmpid.OptionsColumn.AllowEdit = true;
                GridColEmpid.OptionsColumn.AllowFocus = true;
                GridColDept.OptionsColumn.AllowEdit = true;
                GridColDept.OptionsColumn.AllowFocus = true;
            }
            else
            {
                GridColEmpid.OptionsColumn.AllowEdit = false;
                GridColEmpid.OptionsColumn.AllowFocus = false;
                GridColDept.OptionsColumn.AllowEdit = false;
                GridColDept.OptionsColumn.AllowFocus = false;
            }
        }

        private void RepoGridLkupEmpid_EditValueChanged(object sender, EventArgs e)
        {
            GridLookUpEditBase lookUpEdit = (GridLookUpEditBase)sender;
            string sEmpid = lookUpEdit.EditValue?.ToString();
            string sBasdt = DateFrom.EditValue?.ToString().Substring(0, 10);

            if (!string.IsNullOrEmpty(sEmpid))
            {
                DataTable dtRetr = GridRetr.DataSource as DataTable;

                if (dtRetr != null)
                {
                    for (int i = 0; i < dtRetr.Rows.Count; i++)
                    {
                        string sEmpid_1 = dtRetr.Rows[i]["EMPID"]?.ToString();

                        if (sEmpid_1.Equals(sEmpid))
                        {
                            XtraMessageBox.Show("이미 같은 사원의 데이터가 존재합니다.");
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "EMPID", null);
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "DEPT_CD", null);
                            GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColBasDt, null);
                            return;
                        }
                    }
                }

                Dictionary<string, string> dicParams = new Dictionary<string, string>();

                dicParams.Clear();
                dicParams.Add("CMD", "LIST3");
                dicParams.Add("EMPID", sEmpid);

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, "DP_SL001F00", dicParams);

                if (dt != null && dt.Rows.Count != 0)
                {
                    string sDept = dt.Rows[0]["DEPT_CD"]?.ToString();
                    string sGENDER_CD = dt.Rows[0]["GENDER_CD"]?.ToString();

                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "DEPT_CD", sDept);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColGender, sGENDER_CD);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColBasDt, sBasdt);
                }
                else
                {
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, "DEPT_CD", null);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColGender, null);
                    GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridColBasDt, null);
                }
            }
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateFrom.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevDate(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateFrom.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateFrom.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextDate(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateFrom.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }

        private void GE001F00_TextChanged(object sender, EventArgs e)
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