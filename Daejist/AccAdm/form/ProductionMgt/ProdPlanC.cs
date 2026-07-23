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
using DevExpress.XtraPivotGrid;
using ComLib;
using MySql.Data.MySqlClient;
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
*            2. 레이아웃 전체 저장 설정
*           
* 수정일자 : 2021-03-03
* 수정자   : 고혜성
* 수정내용 : (현업요청)
*            1. 업무/길로틴/슈레더/정비의 경우 조회쿼리에서 시작시간 기준 올림으로 쿼리수정
*       
* 수정일자 : 2023-01-13
* 수정자   : 정은영
* ID       : #0001
* 수정내용 : 저장시 pFrm 이 null 이어서 뜨는 오류 수정
*            owner를 지정했으나 안먹혀서 전역변수 선언 후 값 지정
*/
namespace AccAdm
{
    public partial class ProdPlanC : DevExpress.XtraEditors.XtraForm
    {
        public ProdPlanC()
        {
            InitializeComponent();
        }

        public ProdStatus _PRTFRM;//#0001
        public DataRow rowUserInfo { get; set; }
        public string sYmd { get; set; }
        public string sYn = "N";
        public GridView[] arrGrdView;
        private void ProdPlanC_Load(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            DateEditYmd.EditValue = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
            string sWorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            DataTable dtEmpId = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupWorkerID, dtEmpId, GridWorker, GridWorkerNm, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupMntncWorker, dtEmpId, GridMaintenance, GridColMtncChrgNm, "CD", "NM", "");
            
            DataTable dtEquipCd = GetLookUpData("2", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupDeviceNm, dtEquipCd, GridWorker, GridColWorkDeviceName, "CD", "NM", "");

            arrGrdView = new GridView[] { GridViewWorker, GridViewWork, GridViewMaintenance, GridViewCost, GridViewInspect, GirdViewWeek };
            foreach (GridView view in arrGrdView)
            {
                FmMainToolBar2._FontSetting.SetGridView(view);
            }
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }

            GetWorkerInfo(sWorkYmd);

            Cursor = Cursors.Default;


        }

        private void ProdPlanC_Shown(object sender, EventArgs e)
        {
            string sWorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);

            GetWorkAndAprvYNInfo(sWorkYmd);
            GetMaintanenceInfo(sWorkYmd);
            GetCostInfo(sWorkYmd);
            //GetInspectionInfo(sWorkYmd);
        }

        #region[GetLookupData]
        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("N"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine(" UNION ALL");
            }
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT EMP_ID AS CD ");
                strSql.AppendLine("      , EMP_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_NM) AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A");
                strSql.AppendLine("  WHERE EMPL_GB = 'Y' ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.MG_NO AS CD");
                strSql.AppendLine("      , A.EQUIP_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY MG_NO) AS SEQ");
                strSql.AppendLine("   FROM EQUIP_CD A ");
                strSql.AppendLine("  WHERE USE_YN = 'Y' ");
            }

            if (sParam.Equals("Y"))
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
        #endregion[GetLookupData]

        #region[근태 및 결재여부 조회]

        private void GetWorkAndAprvYNInfo(string Ymd)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT MDATE ");
            strSql.AppendLine("      , GUBUN ");
            strSql.AppendLine("      , SIGN1 ");
            strSql.AppendLine("      , SIGN2 ");
            strSql.AppendLine("      , SIGN3 ");
            strSql.AppendLine("      , SIGN4 ");
            strSql.AppendLine("      , SIGN5 ");
            strSql.AppendLine("      , MCLOSED ");
            strSql.AppendLine("      , MLATENESS ");
            strSql.AppendLine("      , MLEAVE ");
            strSql.AppendLine("      , MGOOUT ");
            strSql.AppendLine("      , MEGRESTN ");
            strSql.AppendLine("      , MCONTENT ");
            strSql.AppendLine("   FROM MAKE_S ");
            strSql.AppendLine("  WHERE MDATE = '" + Ymd + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            string sTeamLeader = dt.Rows[0]["SIGN3"]?.ToString();
            string sDeptManager = dt.Rows[0]["SIGN4"]?.ToString();
            string sRep = dt.Rows[0]["SIGN5"]?.ToString();

            if (sRep.Equals("Y"))
            {
                LblRep.Text = "결재";
                LblDeptManager.Text = "결재";
                LblTeamLeader.Text = "결재";

                BtnRepAprv.Text = "결재취소";

                LORepAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LODeptMngrAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LOTmLdrAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else if (sDeptManager.Equals("Y"))
            {
                LblRep.Text = "";
                LblDeptManager.Text = "결재";
                LblTeamLeader.Text = "결재";

                BtnDeptManagerAprv.Text = "결재취소";
                BtnRepAprv.Text = "결재";

                LORepAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LODeptMngrAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LOTmLdrAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else if (sTeamLeader.Equals("Y"))
            {
                LblRep.Text = "";
                LblDeptManager.Text = "";
                LblTeamLeader.Text = "결재";

                BtnTeamLeaderAprv.Text = "결재취소";
                BtnDeptManagerAprv.Text = "결재";

                LORepAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LODeptMngrAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
                LOTmLdrAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
            else if (!sRep.Equals("Y") && !sDeptManager.Equals("Y") && !sTeamLeader.Equals("Y"))
            {
                LblRep.Text = "";
                LblDeptManager.Text = "";
                LblTeamLeader.Text = "";

                BtnTeamLeaderAprv.Text = "결재";
                LORepAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LODeptMngrAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
                LOTmLdrAprv.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }

            //string sWorkerNum = dt.Rows[0]["MCLOSED"]?.ToString();
            string sLateness = dt.Rows[0]["MLATENESS"]?.ToString();
            string sLeave = dt.Rows[0]["MLEAVE"]?.ToString();
            string sVacation = dt.Rows[0]["MGOOUT"]?.ToString();
            string sEgrestn = dt.Rows[0]["MEGRESTN"]?.ToString();

            //TxtProdWorkerNum.EditValue = string.IsNullOrEmpty(sWorkerNum) ? 0 : Convert.ToDouble(sWorkerNum);
            TxtProdLate.EditValue = string.IsNullOrEmpty(sLateness) ? 0 : Convert.ToDouble(sLateness);
            TxtProdGoOut.EditValue = string.IsNullOrEmpty(sLeave) ? 0 : Convert.ToDouble(sLeave);
            TxtProdClose.EditValue = string.IsNullOrEmpty(sVacation) ? 0 : Convert.ToDouble(sVacation);
            TxtEgrestn.EditValue = string.IsNullOrEmpty(sEgrestn) ? 0 : Convert.ToDouble(sEgrestn);

            string sRemark = dt.Rows[0]["MCONTENT"]?.ToString();
            MemoEditContent.EditValue = sRemark;

            Cursor = Cursors.Default;
        }


        #endregion[근태 및 결재여부 조회]

        #region[작업정보 조회]

        private void GetWorkInfo(string Ymd, string EmpID)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LM ");
            strSql.AppendLine("      , A.M1_TIME_F ");
            strSql.AppendLine("      , A.M1_TIME_T ");
            strSql.AppendLine("      , A.M1_CONTENT ");
            strSql.AppendLine("      , A.M1_DEVICE ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , A.ENT_ID ");
            strSql.AppendLine("      , A.MFY_DT ");
            strSql.AppendLine("      , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_1 A ");
            strSql.AppendLine("   LEFT OUTER JOIN MAKE_M B");
            strSql.AppendLine("     ON B.MAKENO = A.MAKENO ");
            strSql.AppendLine("    AND B.MUSER_ID = " + EmpID + " ");
            strSql.AppendLine("  WHERE B.MDATE = " + Ymd + " ");
            strSql.AppendLine("  ORDER BY A.M1_TIME_F ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridWork.DataSource = dt;

            Cursor = Cursors.Default;
        }
        #endregion[작업정보 조회]

        #region[작업자 조회]

        private void GetWorkerInfo(string Ymd)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.EMP_ID ");
            strSql.AppendLine("      , 'N' AS YN ");
            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
            strSql.AppendLine("  WHERE REAL_DUTY_DEPT = '4300' ");
            strSql.AppendLine("    AND EMPL_GB = 'Y' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MUSER_ID ");
            strSql.AppendLine("   FROM MAKE_M A ");
            strSql.AppendLine("  WHERE MDATE = '" + Ymd + "' ");

            DataTable dtYn = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dtYn.Rows.Count; j++)
                {
                    if (dt.Rows[i]["EMP_ID"].ToString().Equals(dtYn.Rows[j]["MUSER_ID"]))
                    {
                        dt.Rows[i]["YN"] = "Y";
                    }
                }
            }

            GridWorker.DataSource = dt;
            TxtProdWorkerNum.EditValue = dt.Rows.Count;
            Cursor = Cursors.Default;
        }
        #endregion[작업자 조회]



       



        #region[정비 조회]
        private void GetMaintanenceInfo(string Ymd)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LM");
            strSql.AppendLine("      , A.M5_TIME_F ");
            strSql.AppendLine("      , A.M5_TIME_T ");
            strSql.AppendLine("      , A.M5_OVERTIME ");
            //strSql.AppendLine("      , A.M5_DEVICE AS NOUSE ");
            strSql.AppendLine("      , C.EQUIP_NM AS M5_DEVICE ");
            strSql.AppendLine("      , A.M5_DEVICE_CD ");
            strSql.AppendLine("      , A.M5_CHRG_ID ");
            //strSql.AppendLine("      , A.M5_CHRG_NM AS NOUSE1 ");
            strSql.AppendLine("      , D.EMP_NM AS M5_CHRG_NM ");
            strSql.AppendLine(" 	 , A.ENT_DT ");
            strSql.AppendLine(" 	 , A.ENT_ID ");
            strSql.AppendLine(" 	 , A.MFY_DT ");
            strSql.AppendLine(" 	 , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_5 A ");
            strSql.AppendLine("   LEFT OUTER JOIN MAKE_M B");
            strSql.AppendLine("     ON A.MAKENO = B.MAKENO ");
            strSql.AppendLine("     LEFT JOIN equip_cd C ");
            strSql.AppendLine("     ON A.M5_DEVICE_CD = C.MG_NO  ");
            strSql.AppendLine("     LEFT JOIN hr_emp_basis D ");
            strSql.AppendLine("     ON A.M5_CHRG_ID = D.EMP_ID  ");
            strSql.AppendLine("  WHERE B.MDATE = '" + Ymd + "' ");
            strSql.AppendLine("  ORDER BY A.M5_TIME_F ");

            DataTable dtMaintenance = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridMaintenance.DataSource = dtMaintenance;

            Cursor = Cursors.Default;
        }
        #endregion[정비 조회]

        #region[비용 조회]
        private void GetCostInfo(string Ymd)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LN");
            strSql.AppendLine("      , A.ECVNAM ");
            strSql.AppendLine("      , A.ECONTENT ");
            strSql.AppendLine("      , SUM(A.EAMT) AS EAMT ");
            strSql.AppendLine("   FROM MAKE_EXPENSE A ");
            strSql.AppendLine("   LEFT OUTER JOIN MAKE_M B");
            strSql.AppendLine("     ON A.MAKENO = B.MAKENO ");
            strSql.AppendLine("  WHERE B.MDATE = '" + Ymd + "' ");
            strSql.AppendLine("    AND A.GUBUN = '1' ");
            strSql.AppendLine("  GROUP BY A.MAKENO, A.MAKENO_LN, A.ECVNAM, A.ECONTENT ");
            strSql.AppendLine("  ORDER BY A.MAKENO, A.MAKENO_LN ");

            DataTable dtCost = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridCost.DataSource = dtCost;

            Cursor = Cursors.Default;
        }
        #endregion[비용 조회]

        #region[설비일상점검 조회]
        private void GetInspectationInfo(string Ymd, string EmpID)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.M7_MTNC_TIME ");
            strSql.AppendLine("      , A.M7_CHECK01_A ");
            strSql.AppendLine("      , A.M7_CHECK01_B ");
            strSql.AppendLine("      , A.M7_CHECK02_A ");
            strSql.AppendLine("      , A.M7_CHECK02_B ");
            strSql.AppendLine("      , A.M7_CHECK03_A ");
            strSql.AppendLine("      , A.M7_CHECK03_B ");
            strSql.AppendLine("      , A.M7_CHECK04_A ");
            strSql.AppendLine("      , A.M7_CHECK04_B ");
            strSql.AppendLine("      , A.M7_CHECK05_A ");
            strSql.AppendLine("      , A.M7_CHECK05_B ");
            strSql.AppendLine("      , A.M7_CHECK06_A ");
            strSql.AppendLine("      , A.M7_CHECK06_B ");
            strSql.AppendLine("      , A.M7_CHECK07_A ");
            strSql.AppendLine("      , A.M7_CHECK07_B ");
            strSql.AppendLine("      , A.M7_CHECK08_A ");
            strSql.AppendLine("      , A.M7_CHECK08_B ");
            strSql.AppendLine("      , A.M7_CHECK09_A ");
            strSql.AppendLine("      , A.M7_CHECK09_B ");
            strSql.AppendLine("      , A.M7_CHECK10_A ");
            strSql.AppendLine("      , A.M7_CHECK10_B ");
            strSql.AppendLine("      , A.M7_CHECK11_A ");
            strSql.AppendLine("      , A.M7_CHECK11_B ");
            strSql.AppendLine("      , A.M7_CHECK12_A ");
            strSql.AppendLine("      , A.M7_CHECK12_B ");
            strSql.AppendLine("      , A.M7_CHECK13_A ");
            strSql.AppendLine("      , A.M7_CHECK13_B ");
            strSql.AppendLine("      , A.M7_CHECK14_A ");
            strSql.AppendLine("      , A.M7_CHECK14_B ");
            strSql.AppendLine("      , A.M7_CHECK15_A ");
            strSql.AppendLine("      , A.M7_CHECK15_B ");
            strSql.AppendLine("      , A.M7_CHECK16_A ");
            strSql.AppendLine("      , A.M7_CHECK16_B ");
            strSql.AppendLine(" 	 , A.ENT_DT ");
            strSql.AppendLine(" 	 , A.ENT_ID ");
            strSql.AppendLine(" 	 , A.MFY_DT ");
            strSql.AppendLine(" 	 , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_7 A ");
            strSql.AppendLine("   LEFT OUTER JOIN MAKE_M B");
            strSql.AppendLine("     ON B.MAKENO = A.MAKENO ");
            strSql.AppendLine("    AND B.MUSER_ID = " + EmpID + " ");
            strSql.AppendLine("  WHERE B.MDATE = " + Ymd + " ");

            DataTable dtInspect = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            DataRow[] rows = dtInspect.Select();

            if (dtInspect.Rows.Count != 0)
            {
                string sInptTime = dtInspect.Rows[0]["M7_MTNC_TIME"]?.ToString();
                if (!string.IsNullOrEmpty(sInptTime) && sInptTime.Length == 6)
                {
                    string sWorkDate = DateEditYmd.EditValue?.ToString().Substring(0, 10);
                    if (!string.IsNullOrEmpty(sWorkDate))
                    {
                        string sTemp = " " + sInptTime.Substring(0, 2) + ":" + sInptTime.Substring(2, 2) + ":" + sInptTime.Substring(4, 2);
                        TimeEditWork.EditValue = sWorkDate + sTemp;
                    }
                }
            }

            DataTable dtInspectCopy = new DataTable();
            dtInspectCopy.TableName = "MAKE_7";

            dtInspectCopy.Columns.Add("MAKENO");
            dtInspectCopy.Columns.Add("INSPECT_ITEM");
            dtInspectCopy.Columns.Add("BREAK_YN");
            dtInspectCopy.Columns.Add("DESC");

            DataTable dtWeekCopy = new DataTable();
            dtInspectCopy.TableName = "MAKE_7";

            dtWeekCopy.Columns.Add("MAKENO");
            dtWeekCopy.Columns.Add("INSPECT_ITEM");
            dtWeekCopy.Columns.Add("BREAK_YN");
            dtWeekCopy.Columns.Add("DESC");

            DataTable dtMonthCopy = new DataTable();
            dtInspectCopy.TableName = "MAKE_7";

            dtMonthCopy.Columns.Add("MAKENO");
            dtMonthCopy.Columns.Add("INSPECT_ITEM");
            dtMonthCopy.Columns.Add("BREAK_YN");
            dtMonthCopy.Columns.Add("DESC");


            DataRow rowCopy;
            string[] arrInspectItem = { "작동유 누유", "라디에이터 누수", "쿨러 누유", "구리스 주입", "엔진오일 점검", "냉각수 점검"
                                      , "벨트 마모, 장력", "컨트롤 밸브", "모터,펌프 이상음", "엔진 이상음", "스윙모터 기어오일", "각 호스라인 누유"
                                      , "각 설치부 크랙", "각 설치부 이상음", "각 핀, 볼트 조임", "괘도장력, 슈판본트"};

            if (dtInspect.Rows.Count > 0)
            {
                for (int i = 1; i <= arrInspectItem.Length; i++)
                {
                    string temp = string.Empty;
                    if (i < 10)
                    {
                        temp = "0" + i.ToString();
                    }
                    else
                    {
                        temp = i.ToString();
                    }

                    if (i <= 7)
                    {
                        rowCopy = dtInspectCopy.NewRow();

                        rowCopy["MAKENO"] = rows[0]["MAKENO"];
                        rowCopy["INSPECT_ITEM"] = arrInspectItem[i - 1].ToString();
                        rowCopy["BREAK_YN"] = rows[0]["M7_CHECK" + temp + "_A"];
                        rowCopy["DESC"] = rows[0]["M7_CHECK" + temp + "_B"];

                        dtInspectCopy.Rows.Add(rowCopy);
                    }
                    if (i > 7 && i <= 14)
                    {
                        rowCopy = dtWeekCopy.NewRow();

                        rowCopy["MAKENO"] = rows[0]["MAKENO"];
                        rowCopy["INSPECT_ITEM"] = arrInspectItem[i - 1].ToString();
                        rowCopy["BREAK_YN"] = rows[0]["M7_CHECK" + temp + "_A"];
                        rowCopy["DESC"] = rows[0]["M7_CHECK" + temp + "_B"];

                        dtWeekCopy.Rows.Add(rowCopy);
                    }
                    if (i > 14 && i <= 16)
                    {
                        rowCopy = dtMonthCopy.NewRow();

                        rowCopy["MAKENO"] = rows[0]["MAKENO"];
                        rowCopy["INSPECT_ITEM"] = arrInspectItem[i - 1].ToString();
                        rowCopy["BREAK_YN"] = rows[0]["M7_CHECK" + temp + "_A"];
                        rowCopy["DESC"] = rows[0]["M7_CHECK" + temp + "_B"];

                        dtMonthCopy.Rows.Add(rowCopy);
                    }
                }
            }
            else
            {
                for (int i = 1; i <= arrInspectItem.Length; i++)
                {
                    string temp = string.Empty;
                    if (i < 10)
                    {
                        temp = "0" + i.ToString();
                    }
                    else
                    {
                        temp = i.ToString();
                    }

                    if (i <= 7)
                    {
                        rowCopy = dtInspectCopy.NewRow();
                        rowCopy["INSPECT_ITEM"] = arrInspectItem[i - 1].ToString();
                        dtInspectCopy.Rows.Add(rowCopy);
                    }
                    if (i > 7 && i <= 14)
                    {
                        rowCopy = dtWeekCopy.NewRow();
                        rowCopy["INSPECT_ITEM"] = arrInspectItem[i - 1].ToString();
                        dtWeekCopy.Rows.Add(rowCopy);
                    }
                    if (i > 14 && i <= 16)
                    {
                        rowCopy = dtMonthCopy.NewRow();
                        rowCopy["INSPECT_ITEM"] = arrInspectItem[i - 1].ToString();
                        dtMonthCopy.Rows.Add(rowCopy);
                    }
                }
            }
            GridInspect.DataSource = dtInspectCopy;
            GridWeek.DataSource = dtWeekCopy;
            GridMonth.DataSource = dtMonthCopy;

            Cursor = Cursors.Default;
        }

        #endregion[설비일상점검 조회]

        private void ProdPlanC_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {

            }
            else if (e.KeyCode == Keys.F1)
            {

            }
            else if (e.KeyCode == Keys.F3)
            {

            }
            else if (e.KeyCode == Keys.F4)
            {
                BtnSaveAndClose_Click(null, null);
            }
            else if (e.KeyCode == Keys.F5)
            {

            }
            else if (e.KeyCode == Keys.F8)
            {
            }
        }

        private void GridViewWorker_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sEmpID = GridViewWorker.GetFocusedRowCellValue("EMP_ID")?.ToString();
            string sWorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            if (!string.IsNullOrEmpty(sEmpID) || !string.IsNullOrEmpty(sWorkYmd))
            {
                GetWorkInfo(sWorkYmd, sEmpID);
                GetInspectationInfo(sWorkYmd, sEmpID);
                //GetGuillotineInfo(sWorkYmd, sEmpID);
                //GetShrederInfo(sWorkYmd, sEmpID);
            }
        }

        private bool GetAprvAuthorityInfo(string sUsrCD, string sDeptCD, string sGradeCd)
        {
            if (string.IsNullOrEmpty(sUsrCD))
            {
                XtraMessageBox.Show("사용자 로그인 정보가 없습니다.");
                return false;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.INSANO ");
            strSql.AppendLine("   FROM ZUSRLST A ");
            strSql.AppendLine("  WHERE USRCD = " + sUsrCD + " ");

            DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dtChk == null)
            {
                XtraMessageBox.Show("사용자의 권한정보가 없습니다.");
                return false;
            }

            string sEmpID = dtChk.Rows[0]["INSANO"]?.ToString();

            if (string.IsNullOrEmpty(sEmpID))
            {
                XtraMessageBox.Show("해당 사용자의 권한관련 인사번호가 없습니다.");
                return false;
            }

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT B.APRV_ATRT_YN ");
            strSql.AppendLine("      , B.DEPT_CD ");
            strSql.AppendLine("      , B.GRADE_CD ");
            strSql.AppendLine("      , B.EMP_ID ");
            strSql.AppendLine("      , B.EMP_NM ");
            strSql.AppendLine("   FROM ZUSRLST A ");
            strSql.AppendLine("   LEFT OUTER JOIN HR_EMP_BASIS B ");
            strSql.AppendLine("     ON B.EMP_ID = A.INSANO ");
            strSql.AppendLine("  WHERE B.GRADE_CD = '" + sGradeCd + "' ");
            strSql.AppendLine("    AND B.DEPT_CD = '" + sDeptCD + "' ");
            strSql.AppendLine("    AND B.EMPL_GB = 'Y' ");
            strSql.AppendLine("    AND B.EMP_ID = " + sEmpID + " ");

            DataTable dtAprvYN = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            if (dtAprvYN.Rows.Count > 0)
            {
                if (dtAprvYN.Rows[0]["APRV_ATRT_YN"].ToString().Equals("Y"))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void BtnTeamLeaderAprv_Click(object sender, EventArgs e)
        {
            string sUsrCD = rowUserInfo["USRCD"]?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CASE APV1_Y WHEN 'Y' THEN 'Y' ELSE 'N' END APV_YN ");
            strSql.AppendLine("   FROM ZPGMAUT A ");
            strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST B  ");
            strSql.AppendLine("     ON A.USRCD = B.USRCD ");
            strSql.AppendLine("  WHERE B.USRCD = " + sUsrCD + " ");
            strSql.AppendLine("    AND A.PGMID = 'ProdStatus' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0)
            {
                string sYn = dt.Rows[0]["APV_YN"]?.ToString();
                if (sYn.Equals("N"))
                {
                    XtraMessageBox.Show("해당 결재에 대한 권한이 없습니다.");
                    return;
                }
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sWorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);

                string sAprvChk = BtnTeamLeaderAprv.Text;

                if (sAprvChk.Equals("결재"))
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE MAKE_S ");
                    strSql.AppendLine("    SET SIGN3 = 'Y' ");
                    strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");
                }
                else if (sAprvChk.Equals("결재취소"))
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE MAKE_S ");
                    strSql.AppendLine("    SET SIGN3 = 'N' ");
                    strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");
                }

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show(sAprvChk + "를 완료했습니다.");

                GetWorkAndAprvYNInfo(sWorkYmd);


                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnDeptManagerAprv_Click(object sender, EventArgs e)
        {
            string sUsrCD = rowUserInfo["USRCD"]?.ToString();
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CASE APV2_Y WHEN 'Y' THEN 'Y' ELSE 'N' END APV_YN ");
            strSql.AppendLine("   FROM ZPGMAUT A ");
            strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST B  ");
            strSql.AppendLine("     ON A.USRCD = B.USRCD ");
            strSql.AppendLine("  WHERE B.USRCD = " + sUsrCD + " ");
            strSql.AppendLine("    AND A.PGMID = 'ProdStatus' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0)
            {
                string sYn = dt.Rows[0]["APV_YN"]?.ToString();
                if (sYn.Equals("N"))
                {
                    XtraMessageBox.Show("해당 결재에 대한 권한이 없습니다.");
                    return;
                }
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sWorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);

                string sAprvChk = BtnDeptManagerAprv.Text;

                if (sAprvChk.Equals("결재"))
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE MAKE_S ");
                    strSql.AppendLine("    SET SIGN4 = 'Y' ");
                    strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");
                }
                else if (sAprvChk.Equals("결재취소"))
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE MAKE_S ");
                    strSql.AppendLine("    SET SIGN4 = 'N' ");
                    strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");
                }

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show(sAprvChk + "를 완료했습니다.");

                GetWorkAndAprvYNInfo(sWorkYmd);


                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnRepAprv_Click(object sender, EventArgs e)
        {
            string sUsrCD = rowUserInfo["USRCD"]?.ToString();
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CASE APV3_Y WHEN 'Y' THEN 'Y' ELSE 'N' END APV_YN ");
            strSql.AppendLine("   FROM ZPGMAUT A ");
            strSql.AppendLine("   LEFT OUTER JOIN ZUSRLST B  ");
            strSql.AppendLine("     ON A.USRCD = B.USRCD ");
            strSql.AppendLine("  WHERE B.USRCD = " + sUsrCD + " ");
            strSql.AppendLine("    AND A.PGMID = 'ProdStatus' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0)
            {
                string sYn = dt.Rows[0]["APV_YN"]?.ToString();
                if (sYn.Equals("N"))
                {
                    XtraMessageBox.Show("해당 결재에 대한 권한이 없습니다.");
                    return;
                }
            }

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sWorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);

                string sAprvChk = BtnRepAprv.Text;

                if (sAprvChk.Equals("결재"))
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE MAKE_S ");
                    strSql.AppendLine("    SET SIGN5 = 'Y' ");
                    strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");
                }
                else if (sAprvChk.Equals("결재취소"))
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE MAKE_S ");
                    strSql.AppendLine("    SET SIGN5 = 'N' ");
                    strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");
                }

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show(sAprvChk + "를 완료했습니다.");

                GetWorkAndAprvYNInfo(sWorkYmd);

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;

                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnSaveAndClose_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder strSql = new StringBuilder();
                string sWorkDate = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT COUNT(A.SIGN1) AS CNT ");
                strSql.AppendLine("      , CASE WHEN ISNULL(MAX(A.SIGN3), '') = 'Y' THEN 'Y' ELSE 'N' END AS YN ");
                strSql.AppendLine("   FROM MAKE_S A ");
                strSql.AppendLine("  WHERE A.MDATE = '" + sWorkDate + "' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt.Rows.Count > 0)
                {
                    string sYn = dt.Rows[0]["YN"].ToString();
                    if (sYn.Equals("Y"))
                    {
                        XtraMessageBox.Show(string.Format("{0}-{1}-{2}의 생산일보내역은 결재처리가 완료되었습니다.", sWorkDate.Substring(0, 4), sWorkDate.Substring(4, 2), sWorkDate.Substring(6, 2)));
                        return;
                    }
                }

                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sClosed = TxtProdWorkerNum.EditValue?.ToString();
                double dClosed = string.IsNullOrEmpty(sClosed) ? 0 : Convert.ToDouble(sClosed);
                string sLateNess = TxtProdLate.EditValue?.ToString();
                double dLateNess = string.IsNullOrEmpty(sLateNess) ? 0 : Convert.ToDouble(sLateNess);
                string sLeave = TxtProdGoOut.EditValue?.ToString();
                double dLeave = string.IsNullOrEmpty(sLeave) ? 0 : Convert.ToDouble(sLeave);
                string sGoOut = TxtProdClose.EditValue?.ToString();
                double dGoOut = string.IsNullOrEmpty(sGoOut) ? 0 : Convert.ToDouble(sGoOut);
                string sContent = MemoEditContent.EditValue?.ToString();

                string sMegrestn = TxtEgrestn.EditValue?.ToString();
                double dMegrestn = string.IsNullOrEmpty(sMegrestn) ? 0 : Convert.ToDouble(sMegrestn);

                strSql.Clear();
                strSql.AppendLine(" ");
                #region mariaDB
                //strSql.AppendLine(" INSERT INTO MAKE_S ");
                //strSql.AppendLine("           ( ");
                //strSql.AppendLine("             MDATE ");
                //strSql.AppendLine("           , GUBUN ");
                //strSql.AppendLine("           , MCLOSED ");
                //strSql.AppendLine("           , MLATENESS ");
                //strSql.AppendLine("           , MLEAVE ");
                //strSql.AppendLine("           , MGOOUT ");
                //strSql.AppendLine("           , MEGRESTN ");
                //strSql.AppendLine("           , MCONTENT ");
                //strSql.AppendLine("           ) ");
                //strSql.AppendLine("           VALUES ");
                //strSql.AppendLine("           ( ");
                //strSql.AppendLine("             '" + sWorkDate + "' ");
                //strSql.AppendLine("           , '1' ");
                //strSql.AppendLine("           , " + dClosed + " ");
                //strSql.AppendLine("           , " + dLateNess + " ");
                //strSql.AppendLine("           , " + dLeave + " ");
                //strSql.AppendLine("           , " + dGoOut + " ");
                //strSql.AppendLine("           , " + dMegrestn + " "); 
                //strSql.AppendLine("           , '" + sContent + "' ");
                //strSql.AppendLine("           ) ");
                //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                //strSql.AppendLine("             MCLOSED = " + dClosed + " ");
                //strSql.AppendLine("           , MLATENESS = " + dLateNess + " ");
                //strSql.AppendLine("           , MLEAVE = " + dLeave + " ");
                //strSql.AppendLine("           , MGOOUT = " + dGoOut + " ");
                //strSql.AppendLine("           , MEGRESTN = " + dMegrestn + " ");
                //strSql.AppendLine("           , MCONTENT = '" + sContent + "' ");
                #endregion

                strSql.AppendLine("IF EXISTS(SELECT * FROM MAKE_S WHERE MDATE = " + sWorkDate + " AND GUBUN = '1')");
                strSql.AppendLine("   BEGIN");
                strSql.AppendLine("         UPDATE MAKE_S");
                strSql.AppendLine("            SET MCLOSED = " + dClosed + " ");
                strSql.AppendLine("              , MLATENESS = " + dLateNess + " ");
                strSql.AppendLine("              , MLEAVE = " + dLeave + " ");
                strSql.AppendLine("              , MGOOUT = " + dGoOut + " ");
                strSql.AppendLine("              , MEGRESTN = " + dMegrestn + " ");
                strSql.AppendLine("              , MCONTENT = '" + sContent + "' ");
                strSql.AppendLine("          WHERE MDATE = " + sWorkDate + " AND GUBUN = '1'");
                strSql.AppendLine("   END");
                strSql.AppendLine("ELSE");
                strSql.AppendLine("   BEGIN");
                strSql.AppendLine("         INSERT INTO MAKE_S ");
                strSql.AppendLine("                   ( ");
                strSql.AppendLine("                     MDATE ");
                strSql.AppendLine("                   , GUBUN ");
                strSql.AppendLine("                   , MCLOSED ");
                strSql.AppendLine("                   , MLATENESS ");
                strSql.AppendLine("                   , MLEAVE ");
                strSql.AppendLine("                   , MGOOUT ");
                strSql.AppendLine("                   , MEGRESTN ");
                strSql.AppendLine("                   , MCONTENT ");
                strSql.AppendLine("                   ) ");
                strSql.AppendLine("                   VALUES ");
                strSql.AppendLine("                   ( ");
                strSql.AppendLine("                     '" + sWorkDate + "' ");
                strSql.AppendLine("                   , '1' ");
                strSql.AppendLine("                   , " + dClosed + " ");
                strSql.AppendLine("                   , " + dLateNess + " ");
                strSql.AppendLine("                   , " + dLeave + " ");
                strSql.AppendLine("                   , " + dGoOut + " ");
                strSql.AppendLine("                   , " + dMegrestn + " ");
                strSql.AppendLine("                   , '" + sContent + "' ");
                strSql.AppendLine("                   ) ");
                strSql.AppendLine("   END");


                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                if (_PRTFRM != null)
                {
                    _PRTFRM.BtnRetr_Click(null, null);
                }

                //Dispose();

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.WaitCursor;

                if (DBConn.dbTran != null)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                }
                Cursor = Cursors.Default;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewWorker_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void GridViewCost_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                string MakeNo = GridViewCost.GetFocusedRowCellValue("MAKENO")?.ToString();
                string Seq = GridViewCost.GetFocusedRowCellValue("MAKENO_LN")?.ToString();
                ProdCostAdder frm = new ProdCostAdder();

                frm.sMakeNo = MakeNo;
                frm.sMakeNoLn = Seq;

                frm.Show();
            }
        }



        private void ProdPlanC_TextChanged(object sender, EventArgs e)
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

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            Dispose();
        }
    }
}