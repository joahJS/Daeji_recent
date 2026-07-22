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
using DevExpress.XtraGrid.Views.Grid;
using MySql.Data.MySqlClient;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraGrid;
using System.Windows.Media;
using System.IO;
using DevExpress.Compression;
using System.Drawing.Imaging;
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
    public partial class ProdMgtReport : DevExpress.XtraEditors.XtraForm
    {
        public ProdMgtReport()
        {
            InitializeComponent();
        }

        public double dMakeNo;
        public string sMakeNo;
        public string ymd;
        public string Firsthhmmss;
        public string sYmd { get; set; }
        public DataRow drUserInfo { get; set; }
        public bool sReadOnly { get; set; }
        public bool MdiYN { get; set; }
        public GridView[] arrGrdView;
        public string ADD_MODIFY_GB;

        private void ProdMgtReport_Load(object sender, EventArgs e)
        {
            sMakeNo = dMakeNo.ToString();
            Cursor = Cursors.WaitCursor;

            ymd = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6 ,2);
            Firsthhmmss = ymd + " " + DateTime.Now.ToString("HH:mm:ss");

            DataTable dtEquipCd = GetLookUpData("3", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupDeviceCd, dtEquipCd, GridWork, GridColWorkDevice, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoMntncGLkupDeviceCd, dtEquipCd, GridMaintenance, GridColMtncDeviceNm, "CD", "NM", "");
            //ComLib.ComGrid.SetGridLookUpEdit(RepoCostGLkupDeviceCd, dtEquipCd, GridCost, GridColCostDeviceNm, "CD", "NM", "");
            DataTable dtEmpID = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoMntncGLkupEmpCd, dtEmpID, GridMaintenance, GridColMtncChrgNm, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupWorkerCd, dtEmpID, GridGuillotine, GridColGilloWorkerCd, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupEmpId, dtEmpID, GridShreder, GridColShrderMgtLine, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupEmpId, dtEmpID, GridShreder, GridColShrderOp, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupEmpId, dtEmpID, GridShreder, GridColShrderCharge, "CD", "NM", "");

            DataTable dtDealer = GetLookUpData("2", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupGumsuDealerCd, dtDealer, GridGumsu, GridColGumsuCVNam, "CD", "NM", "");

            DataTable dtJajae = GetLookUpData("4", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupGradeCd, dtJajae, GridGumsu, GridColGumsuGrade, "CD", "NM", "");

            Cursor = Cursors.Default;

            if (ADD_MODIFY_GB.Equals("ADD"))
            {
                DateEditYmd.EditValue = string.Format("{0}-{1}-{2}", sYmd.Substring(0, 4), sYmd.Substring(4, 2), sYmd.Substring(6, 2));
            }
            else if(ADD_MODIFY_GB.Equals("MOD"))
            {
                DateEditYmd.EditValue = string.Format("{0}-{1}-{2}", sYmd.Substring(0, 4), sYmd.Substring(4, 2), sYmd.Substring(6, 2));
                BtnSave.Enabled = MdiYN;
                BtnCostAdder.Enabled = MdiYN;
            }

            if (!MdiYN)
            {
                GridViewWork.OptionsBehavior.Editable = false;
                GridViewGuillotine.OptionsBehavior.Editable = false;
                GridViewShreder.OptionsBehavior.Editable = false;
                GridViewMaintenance.OptionsBehavior.Editable = false;
                GridViewInspect.OptionsBehavior.Editable = false;
                GridViewGumsu.OptionsBehavior.Editable = false;
            }

            //if (sReadOnly)
            //{
            //    DateEditYmd.ReadOnly = true;
            //}
            //else
            //{
            //    DateEditYmd.ReadOnly = false;
            //    DateEditYmd.EditValue = DateTime.Now;
            //    DateEditYmd.Focus();
            //}
            
            arrGrdView = new GridView[] { GridViewWork, GridViewGuillotine, GridViewShreder, GridViewMaintenance, GridViewCost };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);

            string sFile = ComnEtcFunc.GetLayoutPath() + @"\" + this.Name + "_Layout.xaml";
            if (System.IO.File.Exists(sFile))
            {
                layoutControl1.RestoreLayoutFromXml(sFile);
            }
        }

        private void ProdMgtReport_Shown(object sender, EventArgs e)
        {
            StringBuilder strSql = new StringBuilder();

            string sWorkingDate = sYmd;
            string sEmpID = drUserInfo["EMP_ID"]?.ToString();

            if (string.IsNullOrEmpty(sEmpID))
            {
                XtraMessageBox.Show("사용자 정보가 없습니다.");
                return;
            }

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("   FROM MAKE_M A ");
            strSql.AppendLine("  WHERE MDATE = '" + sWorkingDate + "' ");
            strSql.AppendLine("    AND MUSER_ID = '" + sEmpID + "' ");

            DataTable dtMaster = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            
            if(dtMaster.Rows.Count == 0 || string.IsNullOrEmpty(dtMaster.Rows[0]["MAKENO"].ToString()))
            {
                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT MAX(A.MAKENO) AS MAX_VALUE ");
                strSql.AppendLine("   FROM MAKE_M A ");

                DataTable dtMaxValue = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if(dtMaxValue is null || dtMaxValue.Rows.Count == 0)
                {
                    dMakeNo = 1;
                }
                else
                {
                    string sMaxValue = dtMaxValue.Rows[0]["MAX_VALUE"]?.ToString();
                    dMakeNo = string.IsNullOrEmpty(sMaxValue) ? 1 : Convert.ToDouble(sMaxValue) + 1;
                }
            }
            else if (dtMaster.Rows.Count > 0)
            {
                dMakeNo = Convert.ToDouble(dtMaster.Rows[0]["MAKENO"]);
            }

            Cursor = Cursors.WaitCursor;

            string sMakeNo = dMakeNo.ToString();

            GetMake1Info(sMakeNo);
            GetMake2Info(sMakeNo);
            GetMake3Info(sMakeNo);
            GetMake4Info(sMakeNo);
            GetMake5Info(sMakeNo);
            GetMake6Info(sMakeNo);
            GetMake7Info(sMakeNo);

            //GridViewWork.AddNewRow();
            //GridViewGuillotine.AddNewRow();
            //GridViewShreder.AddNewRow();
            //GridViewMaintenance.AddNewRow();
            //GridViewGumsu.AddNewRow();



            GridViewWork.Focus();
            GridViewWork.FocusedRowHandle = GridViewWork.RowCount - 1;
            GridViewWork.FocusedColumn = GridColWorkStrtTime;

            Cursor = Cursors.Default;
        }
        
        #region[Make_1~7 조회쿼리]

        private void GetMake1Info(string sMakeNo)
        {
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
            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");

            DataTable dtWork = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            DataRow row = dtWork.NewRow();

            row["MAKENO"] = sMakeNo;
            row["MAKENO_LM"] = dtWork.Rows.Count - 1;

            dtWork.Rows.Add(row);

            GridWork.DataSource = dtWork;
        }

        private void GetMake2Info(string sMakeNo)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LM");
            strSql.AppendLine("      , A.M2_TIME_F ");
            strSql.AppendLine("      , A.M2_TIME_T ");
            strSql.AppendLine("      , A.M2_USER ");
            strSql.AppendLine("      , A.M2_USER_ID ");
            strSql.AppendLine("      , A.M2_CHARGE ");
            strSql.AppendLine("      , A.M2_D_PUT ");
            strSql.AppendLine("      , A.M2_D_IN ");
            strSql.AppendLine("      , A.M2_CP_MOVE ");
            strSql.AppendLine("      , A.M2_CP_PUT ");
            strSql.AppendLine("      , A.M2_CP_STOCK ");
            strSql.AppendLine(" 	 , A.ENT_DT ");
            strSql.AppendLine(" 	 , A.ENT_ID ");
            strSql.AppendLine(" 	 , A.MFY_DT ");
            strSql.AppendLine(" 	 , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_2 A");
            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
            strSql.AppendLine("  ORDER BY A.MAKENO_LM ");

            DataTable dtGilo = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            DataRow row = dtGilo.NewRow();

            row["MAKENO"] = sMakeNo;
            row["MAKENO_LM"] = dtGilo.Rows.Count - 1;

            dtGilo.Rows.Add(row);

            GridGuillotine.DataSource = dtGilo;
        }

        private void GetMake3Info(string sMakeNo)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LM");
            strSql.AppendLine("      , A.M3_TIME_F ");
            strSql.AppendLine("      , A.M3_TIME_T ");
            strSql.AppendLine("      , A.M3_TIME ");
            strSql.AppendLine("      , A.M3_IN ");
            strSql.AppendLine("      , A.M3_CP ");
            strSql.AppendLine("      , A.M3_Dmt ");
            strSql.AppendLine("      , A.M3_WGT ");
            strSql.AppendLine("      , A.M3_BAN ");
            strSql.AppendLine("      , A.M3_LIGHT ");
            strSql.AppendLine("      , A.M3_CAR_AVG ");
            strSql.AppendLine("      , A.M3_LINE ");
            strSql.AppendLine("      , A.M3_PUT ");
            strSql.AppendLine("      , A.M3_OP ");
            strSql.AppendLine("      , A.M3_M3_FIND ");
            strSql.AppendLine(" 	 , A.ENT_DT ");
            strSql.AppendLine(" 	 , A.ENT_ID ");
            strSql.AppendLine(" 	 , A.MFY_DT ");
            strSql.AppendLine(" 	 , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_3 A ");
            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
            strSql.AppendLine("  ORDER BY A.MAKENO_LM ");

            DataTable dtShreder = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            DataRow row = dtShreder.NewRow();

            row["MAKENO"] = sMakeNo;
            row["MAKENO_LM"] = dtShreder.Rows.Count - 1;

            dtShreder.Rows.Add(row);
            GridShreder.DataSource = dtShreder;
        }

        private void GetMake4Info(string sMakeNo)
        {
            string sUsrCd = FmMainToolBar2.UserID;
            string sYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.*, C.EMP_NM AS ENT_NM");
            strSql.AppendLine("   FROM MAKE_4 A ");
            strSql.AppendLine("   LEFT JOIN MAKE_M B ");
            strSql.AppendLine("     ON A.MAKENO = B.MAKENO");
            strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS C ");
            strSql.AppendLine("     ON B.MUSER_ID = C.EMP_ID ");
            strSql.AppendLine("  WHERE B.MDATE = '" + sYmd + "' ");
            
            //strSql.Clear();
            //strSql.AppendLine(" SELECT A.MAKENO ");
            //strSql.AppendLine("      , A.MAKENO_LM ");
            //strSql.AppendLine("      , A.M4_CARNO ");
            //strSql.AppendLine("      , A.M4_CVCOD ");
            //strSql.AppendLine("      , A.M4_CVNAM ");
            //strSql.AppendLine("      , A.M4_CVNAM_IDTNO ");
            //strSql.AppendLine("      , A.M4_GRADE ");
            //strSql.AppendLine("      , A.M4_GRADE_CD ");
            //strSql.AppendLine("      , A.M4_WGT ");
            //strSql.AppendLine("      , A.M4_MINUS ");
            //strSql.AppendLine("      , A.M4_EVIDENCE ");
            //strSql.AppendLine("      , A.IMG_CNT ");
            //strSql.AppendLine("      , A.ENT_DT ");
            //strSql.AppendLine("      , A.ENT_ID ");
            //strSql.AppendLine("      , A.MFY_DT ");
            //strSql.AppendLine("      , A.MFY_ID ");
            //strSql.AppendLine("   FROM MAKE_4 A ");
            //strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");

            DataTable dtGumsu = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            //DataRow row = dtGumsu.NewRow();

            //row["MAKENO"] = sMakeNo;
            //row["MAKENO_LM"] = dtGumsu.Rows.Count - 1;

            //dtGumsu.Rows.Add(row);
            GridGumsu.DataSource = dtGumsu;
        }

        private void GetMake5Info(string sMakeNo)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LM");
            strSql.AppendLine("      , A.M5_TIME_F ");
            strSql.AppendLine("      , A.M5_TIME_T ");
            strSql.AppendLine("      , A.M5_OVERTIME ");
            strSql.AppendLine("      , A.M5_DEVICE ");
            strSql.AppendLine("      , A.M5_DEVICE_CD ");
            strSql.AppendLine("      , A.M5_CHRG_ID ");
            strSql.AppendLine("      , A.M5_CHRG_NM ");
            strSql.AppendLine(" 	 , A.ENT_DT ");
            strSql.AppendLine(" 	 , A.ENT_ID ");
            strSql.AppendLine(" 	 , A.MFY_DT ");
            strSql.AppendLine(" 	 , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_5 A ");
            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
            strSql.AppendLine("  ORDER BY A.MAKENO_LM ");

            DataTable dtMaintenance = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            DataRow row = dtMaintenance.NewRow();

            row["MAKENO"] = sMakeNo;
            row["MAKENO_LM"] = dtMaintenance.Rows.Count - 1;

            dtMaintenance.Rows.Add(row);
            GridMaintenance.DataSource = dtMaintenance;
        }

        private void GetMake6Info(string sMakeNo)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LN");
            strSql.AppendLine("      , A.ECVNAM ");
            strSql.AppendLine("      , A.ECONTENT ");
            strSql.AppendLine("      , SUM(A.EAMT) AS EAMT ");
            strSql.AppendLine("   FROM MAKE_EXPENSE A ");
            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
            strSql.AppendLine("  GROUP BY A.MAKENO ");
            strSql.AppendLine(", A.MAKENO_LN");
            strSql.AppendLine(", A.ECVNAM");
            strSql.AppendLine(", A.ECONTENT");
            strSql.AppendLine("  ORDER BY A.MAKENO_LN ");

            DataTable dtCost = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridCost.DataSource = dtCost;
        }

        private void GetMake7Info(string sMakeNo)
        {
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
            strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");

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
                    if(i > 14 && i <= 16)
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
        }

        #endregion[Make_M, 1~7 조회쿼리]

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
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '공용' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.EMP_ID AS CD");
                strSql.AppendLine("      , A.EMP_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_ID) AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A");
                strSql.AppendLine("  WHERE EMPL_GB = 'Y' ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT CAST(A.DEALER_CD AS VARCHAR) AS CD");
                strSql.AppendLine("      , A.DEALER_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_CD) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
                strSql.AppendLine("  WHERE EOB_YN = 'N'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.MG_NO AS CD");
                strSql.AppendLine("      , A.EQUIP_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY MG_NO) AS SEQ");
                strSql.AppendLine("   FROM EQUIP_CD A ");
                strSql.AppendLine("  WHERE USE_YN = 'Y' ");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT CAST(A.J_SERIAL AS VARCHAR) AS CD");
                strSql.AppendLine("      , A.GUBUN1 AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY J_SERIAL) AS SEQ ");
                strSql.AppendLine("   FROM JAJAE A ");
            }

            if (sParam.Equals("Y"))
            {
                strSql.AppendLine(") ");
                strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
                strSql.AppendLine("     ORDER BY SEQ");
            }

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            return dt;
        }

        #endregion[LookUpData Setting]

        #region[ 각 그리드 단축키 설정]

        //GridViewWork(업무)의 마지막 장비명에서의 Enter키 이벤트
        private void RepoGLkupDeviceCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //GridViewWork의 포커스 Index가 마지막 줄일 경우 아래 로직 실행
                if (GridViewWork.FocusedRowHandle == (GridViewWork.RowCount - 1))
                {
                    DataRow row = GridViewWork.GetDataRow(GridViewWork.RowCount - 1);
                    string sM1TimeF = row["M1_TIME_F"].ToString();
                    string sM1TimeT = row["M1_TIME_T"].ToString();
                    if (!string.IsNullOrEmpty(sM1TimeF) && !string.IsNullOrEmpty(sM1TimeT))
                    {
                        DataTable dtWork = (DataTable)GridWork.DataSource;
                        row = dtWork.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dtWork.Rows.Count - 1;
                        dtWork.Rows.Add(row);
                    }
                }
            }
        }

        private void GridWork_ProcessGridKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("추가하시겠습니까?", "업무 항목 추가여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    DataTable dtWork = (DataTable)GridWork.DataSource;
                    DataRow row = dtWork.NewRow();

                    row["MAKENO"] = sMakeNo;
                    row["MAKENO_LM"] = dtWork.Rows.Count - 1;

                    dtWork.Rows.Add(row);

                    GridWork.DataSource = dtWork;
                }
            }
            else if (e.KeyCode == Keys.F4)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("!!! 선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                    , "업무 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    //if (String.IsNullOrEmpty(GridViewGuillotine.GetFocusedRowCellValue("MAKENO").ToString()))
                    //{
                    //    GridView view = sender as GridView;
                    //    view.DeleteRow(view.FocusedRowHandle);
                    //    GridViewGuillotine.RefreshData();
                    //    return;
                    //}

                    ConfirmDelete(GridViewWork, "MAKE_1");
                    GetMake1Info(sMakeNo);
                }
            }
            else if (e.KeyCode == Keys.Down)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                DataTable dt = (DataTable)GridWork.DataSource;

                int iRowHandle = GridViewWork.FocusedRowHandle;
                if (iRowHandle < dt.Rows.Count - 1)
                {
                    return;
                }

                if (dt.Rows.Count > 0)
                {
                    string sM1TimeF = GridViewWork.GetFocusedRowCellValue("M1_TIME_F")?.ToString();
                    string sM1TimeT = GridViewWork.GetFocusedRowCellValue("M1_TIME_T")?.ToString();

                    if (string.IsNullOrEmpty(sM1TimeF) || string.IsNullOrEmpty(sM1TimeT))
                    {
                        return;
                    }
                    else
                    {
                        DataRow row = dt.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dt.Rows.Count - 1;

                        dt.Rows.Add(row);

                        GridWork.DataSource = dt;
                    }
                }
            }

        }

        //길로틴의 마지막 컬럼인 CP재고컬럼 Enter키 시 RowAdd 처리
        private void RepoTxtJaego_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //GridViewWork의 포커스 Index가 마지막 줄일 경우 아래 로직 실행
                if (GridViewGuillotine.FocusedRowHandle == (GridViewGuillotine.RowCount - 1))
                {
                    DataRow row = GridViewGuillotine.GetDataRow(GridViewGuillotine.RowCount - 1);
                    string sM2TimeF = row["M2_TIME_F"].ToString();
                    string sM2TimeT = row["M2_TIME_T"].ToString();
                    if (!string.IsNullOrEmpty(sM2TimeF) && !string.IsNullOrEmpty(sM2TimeT))
                    {
                        DataTable dtGuilo = (DataTable)GridGuillotine.DataSource;
                        row = dtGuilo.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dtGuilo.Rows.Count - 1;
                        dtGuilo.Rows.Add(row);
                    }
                }
            }
        }

        private void GridGuillotine_ProcessGridKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("추가하시겠습니까?", "길로틴 항목 추가여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    DataTable dtGilo = (DataTable)GridGuillotine.DataSource;
                    DataRow row = dtGilo.NewRow();

                    row["MAKENO"] = sMakeNo;
                    row["MAKENO_LM"] = dtGilo.Rows.Count - 1;

                    dtGilo.Rows.Add(row);

                    GridGuillotine.DataSource = dtGilo;
                }
            }

            if (e.KeyCode == Keys.F4)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("!!! 선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                    , "길로틴 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    //if (String.IsNullOrEmpty(GridViewGuillotine.GetFocusedRowCellValue("MAKENO").ToString()))
                    //{
                    //    GridView view = sender as GridView;
                    //    view.DeleteRow(view.FocusedRowHandle);
                    //    GridViewGuillotine.RefreshData();
                    //    return;
                    //}

                    ConfirmDelete(GridViewGuillotine, "MAKE_2");
                    GetMake2Info(sMakeNo);
                }
            }

            if (e.KeyCode == Keys.Down)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                int iRowHandle = GridViewGuillotine.FocusedRowHandle;
                
                DataTable dt = (DataTable)GridGuillotine.DataSource;
                
                if(iRowHandle < dt.Rows.Count - 1)
                {
                    return;
                }

                if (dt.Rows.Count > 0)
                {
                    string sM2TimeF = GridViewGuillotine.GetFocusedRowCellValue("M2_TIME_F")?.ToString();
                    string sM2TimeT = GridViewGuillotine.GetFocusedRowCellValue("M2_TIME_T")?.ToString();

                    if (string.IsNullOrEmpty(sM2TimeF) || string.IsNullOrEmpty(sM2TimeT))
                    {
                        return;
                    }
                    else
                    {
                        DataRow row = dt.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dt.Rows.Count - 1;

                        dt.Rows.Add(row);
                        GridGuillotine.DataSource = dt;
                    }
                }
            }
        }

        //GridShreder 슈레더부분 마지막 컬럼인 비철선별 컬럼에서의 Enter 키다운 시 이벤트
        private void RepoSherderTxt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //GridViewWork의 포커스 Index가 마지막 줄일 경우 아래 로직 실행
                if (GridViewShreder.FocusedRowHandle == (GridViewShreder.RowCount - 1))
                {
                    DataRow row = GridViewShreder.GetDataRow(GridViewShreder.RowCount - 1);
                    string sM3TimeF = row["M3_TIME_F"].ToString();
                    string sM3TimeT = row["M3_TIME_T"].ToString();
                    if (!string.IsNullOrEmpty(sM3TimeF) && !string.IsNullOrEmpty(sM3TimeT))
                    {
                        DataTable dtShreder = (DataTable)GridShreder.DataSource;
                        row = dtShreder.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dtShreder.Rows.Count - 1;
                        dtShreder.Rows.Add(row);
                    }
                }
            }
        }

        private void GridShreder_ProcessGridKey(object sender, KeyEventArgs e)
        {
            if (CheckApprovalYN())
            {
                XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                return;
            }

            if (e.KeyCode == Keys.F1)
            {
                if (XtraMessageBox.Show("추가하시겠습니까?", "슈레더 항목 추가여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    DataTable dtShreder = (DataTable)GridShreder.DataSource;
                    DataRow row = dtShreder.NewRow();

                    row["MAKENO"] = sMakeNo;
                    row["MAKENO_LM"] = dtShreder.Rows.Count - 1;

                    dtShreder.Rows.Add(row);

                    GridShreder.DataSource = dtShreder;
                }
            }

            if (e.KeyCode == Keys.F4)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("!!! 선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                    , "슈레더 선택항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    //if (String.IsNullOrEmpty(GridViewShreder.GetFocusedRowCellValue("MAKENO").ToString()))
                    //{
                    //    GridView view = sender as GridView;
                    //    view.DeleteRow(view.FocusedRowHandle);
                    //    GridViewShreder.RefreshData();
                    //    return;
                    //}

                    ConfirmDelete(GridViewShreder, "MAKE_3");
                    GetMake3Info(sMakeNo);
                }
            }

            if(e.KeyCode == Keys.Down)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                DataTable dt = (DataTable)GridShreder.DataSource;

                int iRowHandle = GridViewShreder.FocusedRowHandle;
                if (iRowHandle < dt.Rows.Count - 1)
                {
                    return;
                }

                if (dt.Rows.Count > 0)
                {
                    string sM3TimeF = GridViewShreder.GetFocusedRowCellValue("M3_TIME_F")?.ToString();
                    string sM3TimeT = GridViewShreder.GetFocusedRowCellValue("M3_TIME_T")?.ToString();

                    if (string.IsNullOrEmpty(sM3TimeF) || string.IsNullOrEmpty(sM3TimeT))
                    {
                        return;
                    }
                    else
                    {
                        DataRow row = dt.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dt.Rows.Count - 1;

                        dt.Rows.Add(row);
                        GridShreder.DataSource = dt;
                    }
                }
            }
        }

        //GridViewMaintenance 정비그리드 부분 마지막 컬럼인 담당자 컬럼에 대한 Enter키 처리
        private void RepoMntncGLkupEmpCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //GridViewWork의 포커스 Index가 마지막 줄일 경우 아래 로직 실행
                if (GridViewMaintenance.FocusedRowHandle == (GridViewMaintenance.RowCount - 1))
                {
                    DataRow row = GridViewMaintenance.GetDataRow(GridViewMaintenance.RowCount - 1);
                    string sM5TimeF = row["M5_TIME_F"].ToString();
                    string sM5TimeT = row["M5_TIME_T"].ToString();
                    
                    if (!string.IsNullOrEmpty(sM5TimeF) && !string.IsNullOrEmpty(sM5TimeT))
                    {
                        DataTable dtMaintenance = (DataTable)GridMaintenance.DataSource;
                        row = dtMaintenance.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dtMaintenance.Rows.Count - 1;
                        dtMaintenance.Rows.Add(row);
                    }
                }
            }
        }

        private void GridMaintenance_ProcessGridKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("추가하시겠습니까?", "정비 항목 추가여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    DataTable dtMaintenance = (DataTable)GridMaintenance.DataSource;
                    DataRow row = dtMaintenance.NewRow();

                    row["MAKENO"] = sMakeNo;
                    row["MAKENO_LM"] = dtMaintenance.Rows.Count - 1;

                    dtMaintenance.Rows.Add(row);

                    GridMaintenance.DataSource = dtMaintenance;
                }
            }

            if (e.KeyCode == Keys.F4)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("!!! 선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                    , "정비 선택항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    //if (String.IsNullOrEmpty(GridViewMaintenance.GetFocusedRowCellValue("MAKENO").ToString()))
                    //{
                    //    GridView view = sender as GridView;
                    //    view.DeleteRow(view.FocusedRowHandle);
                    //    GridViewMaintenance.RefreshData();
                    //    return;
                    //}

                    ConfirmDelete(GridViewMaintenance, "MAKE_5");
                    GetMake5Info(sMakeNo);
                }
            }

            if (e.KeyCode == Keys.Down)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                DataTable dt = (DataTable)GridMaintenance.DataSource;

                int iRowHandle = GridViewMaintenance.FocusedRowHandle;
                if (iRowHandle < dt.Rows.Count - 1)
                {
                    return;
                }

                if (dt.Rows.Count > 0)
                {
                    string sM5TimeF = GridViewMaintenance.GetFocusedRowCellValue("M5_TIME_F")?.ToString();
                    string sM5TimeT = GridViewMaintenance.GetFocusedRowCellValue("M5_TIME_T")?.ToString();

                    if (string.IsNullOrEmpty(sM5TimeF) || string.IsNullOrEmpty(sM5TimeT))
                    {
                        return;
                    }
                    else
                    {
                        DataRow row = dt.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dt.Rows.Count - 1;

                        dt.Rows.Add(row);
                        GridMaintenance.DataSource = dt;
                    }
                }
            }
        }

        private void GridCost_ProcessGridKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F4)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("!!! 선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                    , "비용 선택항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    try
                    {
                        string sMakeNo = GridViewCost.GetFocusedRowCellValue("MAKENO").ToString();
                        string sMakeNoLm = GridViewCost.GetFocusedRowCellValue("MAKENO_LN").ToString();

                        DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                        SqlCommand cmd = DBConn.dbCon.CreateCommand();
                        cmd.Transaction = DBConn.dbTran;

                        StringBuilder strSql = new StringBuilder();

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" DELETE ");
                        strSql.AppendLine("   FROM EQUIP_CD_HISTORY ");
                        strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
                        strSql.AppendLine("    AND MAKENO_LN = " + sMakeNoLm + " ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" DELETE ");
                        strSql.AppendLine("   FROM MAKE_EXPENSE ");
                        strSql.AppendLine("  WHERE MAKENO = " + sMakeNo + " ");
                        strSql.AppendLine("    AND MAKENO_LN = " + sMakeNoLm + " ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        DBConn.dbTran.Commit();
                        DBConn.dbTran = null;

                        XtraMessageBox.Show("삭제가 완료되었습니다.");
                    }
                    catch (Exception ex)
                    {
                        Cursor = Cursors.Default;

                        DBConn.dbTran.Rollback();
                        DBConn.dbTran = null;
                        XtraMessageBox.Show(ex.Message);
                    }

                    GetMake6Info(sMakeNo);
                }
            }
        }

        private void GridGumsu_ProcessGridKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("추가하시겠습니까?", "검수 항목 추가여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    DataTable dtGumsu = (DataTable)GridGumsu.DataSource;
                    DataRow row = dtGumsu.NewRow();

                    row["MAKENO"] = sMakeNo;
                    row["MAKENO_LM"] = dtGumsu.Rows.Count - 1;

                    dtGumsu.Rows.Add(row);

                    GridGumsu.DataSource = dtGumsu;
                }
            }

            if (e.KeyCode == Keys.F4)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("!!! 선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                    , "검수 선택항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    //if (String.IsNullOrEmpty(GridViewGumsu.GetFocusedRowCellValue("MAKENO").ToString()))
                    //{
                    //    GridView view = sender as GridView;
                    //    view.DeleteRow(view.FocusedRowHandle);
                    //    GridViewGumsu.RefreshData();
                    //    return;
                    //}

                    ConfirmDelete(GridViewGumsu, "MAKE_6");
                    GetMake6Info(sMakeNo);
                }
            }

            if (e.KeyCode == Keys.Down)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                DataTable dt = (DataTable)GridGumsu.DataSource;

                int iRowHandle = GridViewGumsu.FocusedRowHandle;
                if (iRowHandle < dt.Rows.Count - 1)
                {
                    return;
                }

                if (dt.Rows.Count > 0)
                {
                    string sDealer = GridViewGumsu.GetFocusedRowCellValue("M4_CVNAM")?.ToString();

                    if (string.IsNullOrEmpty(sDealer))
                    {
                        return;
                    }
                    else
                    {
                        DataRow row = dt.NewRow();

                        row["MAKENO"] = sMakeNo;
                        row["MAKENO_LM"] = dt.Rows.Count - 1;

                        dt.Rows.Add(row);
                        GridGumsu.DataSource = dt;
                    }
                }
            }
        }

        private void ConfirmDelete(GridView gridView, string sTableNm)
        {
            try
            {
                string sMakeNo = gridView.GetFocusedRowCellValue("MAKENO").ToString();
                string sMakeNoLm = gridView.GetFocusedRowCellValue("MAKENO_LM").ToString();

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" DELETE ");
                strSql.AppendLine("   FROM " + sTableNm + "");
                strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");
                strSql.AppendLine("    AND MAKENO_LM = " + sMakeNoLm + " ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                XtraMessageBox.Show("삭제가 완료되었습니다.");
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        #endregion[각 그리드 단축키 설정]
        
        #region[RepositoryItem 관련 Events]

        private void RepoGumsuBtnEdit_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PopUpDealerCd frm = new PopUpDealerCd();

            DataRow drDealerInfo;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                drDealerInfo = frm.drResult;
                (sender as ButtonEdit).Text = drDealerInfo["DEALER_NM"].ToString();
                GridViewGumsu.SetFocusedRowCellValue("M4_CVNAM_IDTNO", drDealerInfo["IDT_NO"].ToString());
            }
        }

        private void RepoCostBtnEdit_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PopUpDealerCd frm = new PopUpDealerCd();

            DataRow drDealerInfo;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                drDealerInfo = frm.drResult;
                (sender as ButtonEdit).Text = drDealerInfo["DEALER_NM"].ToString();
                GridViewCost.SetFocusedRowCellValue("M6_CVNAM_CD", drDealerInfo["DEALER_CD"].ToString());
            }
        }

        private void RepoCostBtnEditDevice_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PopUpToolCd frm = new PopUpToolCd();

            DataRow drToolInfo;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                drToolInfo = frm.drResult;
                (sender as ButtonEdit).Text = drToolInfo["TOOL_NM"].ToString();
                GridViewCost.SetFocusedRowCellValue("M6_DEVICE_CD", drToolInfo["MG_NO"].ToString());
            }
        }

        private void RepoMntncGLkupDeviceCd_Leave(object sender, EventArgs e)
        {
            string sDeviceNm = GridViewMaintenance.GetFocusedRowCellDisplayText("M5_DEVICE_CD");
            GridViewMaintenance.SetFocusedRowCellValue("M5_DEVICE", sDeviceNm);
        }

        private void RepoMntncGLkupEmpCd_Leave(object sender, EventArgs e)
        {
            string sEmpNm = GridViewMaintenance.GetFocusedRowCellDisplayText("M5_CHRG_ID");
            GridViewMaintenance.SetFocusedRowCellValue("M5_CHRG_NM", sEmpNm);
        }

        private void RepoCostGLkupDeviceCd_Leave(object sender, EventArgs e)
        {
            string sDeviceNm = GridViewMaintenance.GetFocusedRowCellDisplayText("M6_DEVICE_CD");
            GridViewMaintenance.SetFocusedRowCellValue("M6_DEVICE", sDeviceNm);
        }

        private void RepoGLkupGradeCd_Leave(object sender, EventArgs e)
        {
            string sGradeCd = GridViewGumsu.GetFocusedRowCellDisplayText("M4_GRADE_CD");
            GridViewGumsu.SetFocusedRowCellValue("M4_GRADE", sGradeCd);
        }

        private void RepoGLkupWorkerCd_Leave(object sender, EventArgs e)
        {
            string sEmpNm = GridViewGuillotine.GetFocusedRowCellDisplayText("M2_USER_ID");
            GridViewGuillotine.SetFocusedRowCellValue("M2_USER", sEmpNm);
        }
        #endregion[RepositoryItem 관련 Events]

        private void BtnSave_Click(object sender, EventArgs e)
        {
            string sWorkYmd1 = DateEditYmd.EditValue?.ToString().Substring(0, 10);
            string sWorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sEntID = FmMainToolBar2.UserID;
            string sEmpID = drUserInfo["EMP_ID"].ToString();
            string sEmpNm = drUserInfo["EMP_NM"].ToString();

            if (string.IsNullOrEmpty(sWorkYmd))
            {
                XtraMessageBox.Show("일자가 선택되지 않았습니다.");
                return;
            }

            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.REAL_DUTY_DEPT ");
            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpID + "' ");
            strSql.AppendLine("    AND EMPL_GB = 'Y' ");
            DataTable dtDept = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            string sDept = dtDept.Rows[0]["REAL_DUTY_DEPT"]?.ToString();


            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.MDATE ");
            if (sDept == "4100") { strSql.AppendLine("       ,A.SIGN1  AS Sign"); }
            if (sDept == "4150") { strSql.AppendLine("       ,A.SIGN1a  AS Sign"); }
            if (sDept == "4200") { strSql.AppendLine("       ,A.SIGN2  AS Sign"); }
            if (sDept == "4300") { strSql.AppendLine("       ,A.SIGN3  AS Sign"); }
            strSql.AppendLine("   FROM MAKE_S A ");
            strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");

            DataTable MakeSChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(MakeSChk.Rows.Count > 0 && MakeSChk.Rows[0]["Sign"].ToString().Equals("Y"))
            {
                XtraMessageBox.Show(DateEditYmd.EditValue.ToString().Substring(0, 10) + "의 작업정보건은 결제승인완료 상태이므로 수정할 수 없습니다.");
                return;
            }


            //if (!sReadOnly)
            //{
            //    strSql.Clear();
            //    strSql.AppendLine(" ");
            //    strSql.AppendLine(" SELECT A.MDATE ");
            //    strSql.AppendLine("      , A.MUSER_ID ");
            //    strSql.AppendLine("   FROM MAKE_M A ");
            //    strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "' ");
            //    strSql.AppendLine("    AND MUSER_ID = '" + sEmpID + "' ");

            //    DataTable dtCk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //    if (dtCk.Rows.Count > 0)
            //    {
            //        XtraMessageBox.Show(sWorkYmd1 + "의 생산일보 내역이 존재합니다.");
            //        DateEditYmd.Focus();
            //        return;
            //    }
            //}

            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;
                
                Cursor = Cursors.Default;

                #region[MAKE_S 체크 및 Insert]

                string sWorkingYmd = DateEditYmd.EditValue.ToString().Replace("-", "").Substring(0, 8);

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.MDATE ");
                strSql.AppendLine("   FROM MAKE_S A ");
                strSql.AppendLine("  WHERE MDATE = '" + sWorkingYmd + "' ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt.Rows.Count == 0)
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" INSERT INTO MAKE_S ");
                    strSql.AppendLine("           ( MDATE ");
                    strSql.AppendLine("           , GUBUN ");
                    strSql.AppendLine("           , SIGN1 ");
                    strSql.AppendLine("           , SIGN1a ");
                    strSql.AppendLine("           , SIGN2 ");
                    strSql.AppendLine("           , SIGN3 ");
                    strSql.AppendLine("           , MCLOSED ");
                    strSql.AppendLine("           , MLATENESS ");
                    strSql.AppendLine("           , MLEAVE ");
                    strSql.AppendLine("           , MGOOUT ");
                    strSql.AppendLine("           ) ");
                    strSql.AppendLine("      VALUES ");
                    strSql.AppendLine("           ( '" + sWorkingYmd + "' ");
                    strSql.AppendLine("           , '1' ");
                    strSql.AppendLine("           , 'N' ");
                    strSql.AppendLine("           , 'N' ");
                    strSql.AppendLine("           , 'N'");
                    strSql.AppendLine("           , 'N'");
                    strSql.AppendLine("           , 0 ");
                    strSql.AppendLine("           , 0 ");
                    strSql.AppendLine("           , 0 ");
                    strSql.AppendLine("           , 0 ");
                    strSql.AppendLine("           ) ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                #endregion[MAKE_S 체크 및 Insert]
                
                #region[MAKE_M 등록 부분]

                //strSql.Clear();
                //strSql.AppendLine(" ");
                //strSql.AppendLine(" SELECT A.MDATE ");
                //strSql.AppendLine("   FROM MAKE_M A ");
                //strSql.AppendLine("  WHERE MAKENO = " + dMakeNo + "");
                //strSql.AppendLine("    AND MUSER_ID = " + sEmpID + "");

                //DataTable dtMakeM = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                //if(dtMakeM.Rows.Count == 0)
                //{


                    //string sEmpID = FmMainToolBar2.UserID;

                    //if (dMakeNo == 0)
                    //{
                    //    strSql.Clear();
                    //    strSql.AppendLine(" SELECT CASE WHEN MAX(A.MAKENO) IS NULL THEN CONCAT(DATE_FORMAT(NOW(), '%Y%m'), '001')  ");
                    //    strSql.AppendLine("             WHEN MAX(A.MAKENO) IS NOT NULL THEN MAX(A.MAKENO) + 1 END AS MAX_VALUE ");
                    //    strSql.AppendLine("   FROM MAKE_M A ");
                    //    strSql.AppendLine("  WHERE LEFT(A.MAKENO, 6) = DATE_FORMAT(NOW(), '%Y%m')  ");

                    //    DataTable dtMax = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                    //    string sMax = dtMax?.Rows[0]["MAX_VALUE"]?.ToString();

                    //    dMakeNo = Convert.ToDouble(sMax);
                    //}

                    //strSql.Clear();
                    //strSql.AppendLine(" ");
                    //strSql.AppendLine(" INSERT INTO MAKE_M ");
                    //strSql.AppendLine("           ( ");
                    //strSql.AppendLine("             MAKENO ");
                    //strSql.AppendLine("           , MDATE ");
                    //strSql.AppendLine("           , MUSER ");
                    //strSql.AppendLine("           , MUSER_ID ");
                    //strSql.AppendLine("           , ENT_DT ");
                    //strSql.AppendLine("           , ENT_ID ");
                    //strSql.AppendLine("           , MFY_DT ");
                    //strSql.AppendLine("           , MFY_ID ");
                    //strSql.AppendLine("           ) ");
                    //strSql.AppendLine("      VALUES ");
                    //strSql.AppendLine("           ( ");
                    //strSql.AppendLine("             " + dMakeNo + " ");
                    //strSql.AppendLine("           , '" + sWorkingYmd + "' ");
                    //strSql.AppendLine("           , '" + sEmpNm + "' ");
                    //strSql.AppendLine("           , '" + sEmpID + "' ");
                    //strSql.AppendLine("           , NOW() ");
                    //strSql.AppendLine("           , '" + sEntID + "' ");
                    //strSql.AppendLine("           , NOW()");
                    //strSql.AppendLine("           , '" + sEntID + "' ");
                    //strSql.AppendLine("           ) ");
                    //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                    //strSql.AppendLine("             MDATE = '" + sWorkingYmd + "' ");
                    //strSql.AppendLine("           , MUSER = '" + sEmpNm + "' ");
                    //strSql.AppendLine("           , MUSER_ID = '" + sEmpID + "' ");
                    //strSql.AppendLine("           , MFY_DT = NOW() ");
                    //strSql.AppendLine("           , MFY_ID = '" + sEntID + "' ");

                    //cmd.CommandType = CommandType.Text;
                    //cmd.CommandText = strSql.ToString();
                    //cmd.ExecuteNonQuery();

                //}

                #endregion[MAKE_M 등록 부분]

                #region[업무 저장부분]

                if (GridViewWork.RowCount > 0)
                {
                    Cursor = Cursors.WaitCursor;

                    DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridWork.DataSource);
                    DataTable dtMerge = dsSave.Tables[0];

                    for (int i = 0; i < dtMerge.Rows.Count; i++)
                    {
                        if(i == 0 || i == dtMerge.Rows.Count - 1)
                        {
                            string chksM1TimeF = dtMerge.Rows[i]["M1_TIME_F"]?.ToString();
                            string chksM1TimeT = dtMerge.Rows[i]["M1_TIME_F"]?.ToString();
                            string chksM1Content = dtMerge.Rows[i]["M1_CONTENT"]?.ToString();
                            string chksM1Device = dtMerge.Rows[i]["M1_DEVICE"]?.ToString();

                            if(string.IsNullOrEmpty(chksM1TimeF) & string.IsNullOrEmpty(chksM1TimeT)
                              & string.IsNullOrEmpty(chksM1Content) & string.IsNullOrEmpty(chksM1Device))
                            {
                                continue;
                            }
                        }

                        string sMakeNo = dMakeNo.ToString();
                        double dMakeNoLm = 0;

                        if (String.IsNullOrEmpty(dtMerge.Rows[i]["MAKENO_LM"].ToString()))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" SELECT MAX(A.MAKENO_LM)");
                            strSql.AppendLine("   FROM MAKE_1 A ");
                            strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();

                            if (String.IsNullOrEmpty(cmd.ExecuteScalar().ToString()))
                            {
                                dMakeNoLm = 1;
                            }
                            else
                            {
                                dMakeNoLm = Convert.ToDouble(cmd.ExecuteScalar()) + 1;
                            }
                        }
                        else
                        {
                            dMakeNoLm = Convert.ToDouble(dtMerge.Rows[i]["MAKENO_LM"]);
                        }

                        #region 2021-08-31 이전
                        //if (dtMerge.Rows[i]["M1_TIME_F"].ToString().Length > 5)
                        //{
                        //    sM1TimeF = dtMerge.Rows[i]["M1_TIME_F"].ToString().Substring(11, 5);
                        //}
                        //else if (dtMerge.Rows[i]["M1_TIME_F"].ToString().Length == 5)
                        //{
                        //    sM1TimeF = dtMerge.Rows[i]["M1_TIME_F"].ToString();
                        //}

                        //string sM1TimeT = string.Empty;
                        //if (dtMerge.Rows[i]["M1_TIME_T"].ToString().Length > 5)
                        //{
                        //    sM1TimeT = dtMerge.Rows[i]["M1_TIME_T"].ToString().Substring(11, 5);
                        //}
                        //else if (dtMerge.Rows[i]["M1_TIME_T"].ToString().Length == 5)
                        //{
                        //    sM1TimeT = dtMerge.Rows[i]["M1_TIME_T"].ToString();
                        //}
                        #endregion

                        string sM1TimeF = dtMerge.Rows[i]["M1_TIME_F"].ToString();
                        if (DateTime.TryParse(sM1TimeF, out DateTime dtM1TimeF))
                            sM1TimeF = dtM1TimeF.ToString("HH:mm");
                        else
                            sM1TimeF = "";

                        string sM1TimeT = dtMerge.Rows[i]["M1_TIME_T"].ToString();
                        if (DateTime.TryParse(sM1TimeT, out DateTime dtM1TimeT))
                            sM1TimeT = dtM1TimeT.ToString("HH:mm");
                        else
                            sM1TimeT = "";

                        string sM1Content = dtMerge.Rows[i]["M1_CONTENT"].ToString();
                        string sM1Device = dtMerge.Rows[i]["M1_DEVICE"].ToString();
                        string sEntId = FmMainToolBar2.UserID;
                        string sMfyId = FmMainToolBar2.UserID;

                        strSql.Clear();
                        #region mariaDB
                        //strSql.AppendLine(" ");
                        //strSql.AppendLine(" INSERT INTO MAKE_1 ");
                        //strSql.AppendLine("           ( MAKENO ");
                        //strSql.AppendLine("           , MAKENO_LM ");
                        //strSql.AppendLine("           , M1_TIME_F ");
                        //strSql.AppendLine("           , M1_TIME_T ");
                        //strSql.AppendLine("           , M1_CONTENT ");
                        //strSql.AppendLine("           , M1_DEVICE ");
                        //strSql.AppendLine("           , ENT_DT ");
                        //strSql.AppendLine("           , ENT_ID ");
                        //strSql.AppendLine("           , MFY_DT ");
                        //strSql.AppendLine("           , MFY_ID ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("      VALUES ");
                        //strSql.AppendLine("           ( " + sMakeNo + " ");
                        //strSql.AppendLine("           , " + dMakeNoLm + "");
                        //strSql.AppendLine("           , '" + sM1TimeF + "' ");
                        //strSql.AppendLine("           , '" + sM1TimeT + "' ");
                        //strSql.AppendLine("           , '" + sM1Content + "' ");
                        //strSql.AppendLine("           , '" + sM1Device + "' ");
                        //strSql.AppendLine("           , CONVERT([varchar](20),getdate(),(21)) ");
                        //strSql.AppendLine("           , '" + sEntId + "' ");
                        //strSql.AppendLine("           , CONVERT([varchar](20),getdate(),(21)) ");
                        //strSql.AppendLine("           , '" + sMfyId + "') ");
                        //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                        //strSql.AppendLine("             M1_TIME_F = '" + sM1TimeF + "' ");
                        //strSql.AppendLine("           , M1_TIME_T = '" + sM1TimeT + "' ");
                        //strSql.AppendLine("           , M1_CONTENT = '" + sM1Content + "' ");
                        //strSql.AppendLine("           , M1_DEVICE = '" + sM1Device + "' ");
                        //strSql.AppendLine("           , MFY_DT = CONVERT([varchar](20),getdate(),(21)) ");
                        //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                        #endregion

                        strSql.AppendLine("IF EXISTS(SELECT * FROM MAKE_1 WHERE MAKENO = " + sMakeNo + "  AND MAKENO_LM = " + dMakeNoLm + ")");
                        strSql.AppendLine("    BEGIN");
                        strSql.AppendLine("          UPDATE MAKE_1");
                        strSql.AppendLine("             SET M1_TIME_F = '" + sM1TimeF + "' ");
                        strSql.AppendLine("               , M1_TIME_T = '" + sM1TimeT + "' ");
                        strSql.AppendLine("               , M1_CONTENT = '" + sM1Content + "' ");
                        strSql.AppendLine("               , M1_DEVICE = '" + sM1Device + "' ");
                        strSql.AppendLine("               , MFY_DT = CONVERT([varchar](20),getdate(),(21)) ");
                        strSql.AppendLine("               , MFY_ID = '" + sMfyId + "' ");
                        strSql.AppendLine("           WHERE MAKENO = " + sMakeNo + "  AND MAKENO_LM = " + dMakeNoLm);
                        strSql.AppendLine("    END");
                        strSql.AppendLine("    ELSE");
                        strSql.AppendLine("    BEGIN");
                        strSql.AppendLine("          INSERT INTO MAKE_1");
                        strSql.AppendLine("               (MAKENO");
                        strSql.AppendLine("               , MAKENO_LM");
                        strSql.AppendLine("               , M1_TIME_F");
                        strSql.AppendLine("               , M1_TIME_T");
                        strSql.AppendLine("               , M1_CONTENT");
                        strSql.AppendLine("               , M1_DEVICE");
                        strSql.AppendLine("               , ENT_DT");
                        strSql.AppendLine("               , ENT_ID");
                        strSql.AppendLine("               , MFY_DT");
                        strSql.AppendLine("               , MFY_ID )");
                        strSql.AppendLine("          VALUES ");
                        strSql.AppendLine("               ( " + sMakeNo + " ");
                        strSql.AppendLine("               , " + dMakeNoLm + "");
                        strSql.AppendLine("               , '" + sM1TimeF + "' ");
                        strSql.AppendLine("               , '" + sM1TimeT + "' ");
                        strSql.AppendLine("               , '" + sM1Content + "' ");
                        strSql.AppendLine("               , '" + sM1Device + "' ");
                        strSql.AppendLine("               , CONVERT([varchar](20),getdate(),(21)) ");
                        strSql.AppendLine("               , '" + sEntId + "' ");
                        strSql.AppendLine("               , CONVERT([varchar](20),getdate(),(21)) ");
                        strSql.AppendLine("               , '" + sMfyId + "') ");
                        strSql.AppendLine("      END");
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                    }
                    Cursor = Cursors.Default;
                }



                #endregion[업무 저장부분]

                #region[길로틴 저장부분]

                if (GridViewGuillotine.RowCount > 0)
                {
                    Cursor = Cursors.WaitCursor;

                    DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridGuillotine.DataSource);
                    DataTable dtMerge = dsSave.Tables[0];

                    for (int i = 0; i < dtMerge.Rows.Count; i++)
                    {
                        if (i == 0 || i == dtMerge.Rows.Count - 1)
                        {
                            string chksM2TimeF = dtMerge.Rows[i]["M2_TIME_F"]?.ToString();
                            string chksM2TimeT = dtMerge.Rows[i]["M2_TIME_F"]?.ToString();
                            string chksM2User = dtMerge.Rows[i]["M2_USER"].ToString();
                            string chksM2UserId = dtMerge.Rows[i]["M2_USER_ID"].ToString();
                            string chksM2Charge = dtMerge.Rows[i]["M2_CHARGE"].ToString();
                            string chksM2DPut = dtMerge.Rows[i]["M2_D_PUT"].ToString();
                            string chksM2DIn = dtMerge.Rows[i]["M2_D_IN"].ToString();
                            string chksM2CpMove = dtMerge.Rows[i]["M2_CP_MOVE"].ToString();
                            string chksM2CpPut = dtMerge.Rows[i]["M2_CP_PUT"].ToString();
                            string chksM2CpStock = dtMerge.Rows[i]["M2_CP_STOCK"].ToString();
                            
                            if (string.IsNullOrEmpty(chksM2TimeF) & string.IsNullOrEmpty(chksM2TimeT)
                              & string.IsNullOrEmpty(chksM2User) & string.IsNullOrEmpty(chksM2UserId)
                              & string.IsNullOrEmpty(chksM2Charge) & string.IsNullOrEmpty(chksM2DPut)
                              & string.IsNullOrEmpty(chksM2DIn) & string.IsNullOrEmpty(chksM2CpMove)
                              & string.IsNullOrEmpty(chksM2CpPut) & string.IsNullOrEmpty(chksM2CpStock))
                            {
                                continue;
                            }
                        }

                        string sMakeNo = dMakeNo.ToString();
                        double dMakeNoLm = 0;

                        if (String.IsNullOrEmpty(dtMerge.Rows[i]["MAKENO_LM"].ToString()))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" SELECT MAX(A.MAKENO_LM)");
                            strSql.AppendLine("   FROM MAKE_2 A ");
                            strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();

                            if (String.IsNullOrEmpty(cmd.ExecuteScalar().ToString()))
                            {
                                dMakeNoLm = 1;
                            }
                            else
                            {
                                dMakeNoLm = Convert.ToDouble(cmd.ExecuteScalar()) + 1;
                            }
                        }
                        else
                        {
                            dMakeNoLm = Convert.ToDouble(dtMerge.Rows[i]["MAKENO_LM"]);
                        }

                        #region 2021-08-31 이전코드
                        //string sM2TimeF = string.Empty;
                        //if (dtMerge.Rows[i]["M2_TIME_F"].ToString().Length > 5)
                        //{
                        //    sM2TimeF = dtMerge.Rows[i]["M2_TIME_F"].ToString().Substring(11, 5);
                        //}
                        //else if(dtMerge.Rows[i]["M2_TIME_F"].ToString().Length == 5)
                        //{
                        //    sM2TimeF = dtMerge.Rows[i]["M2_TIME_F"].ToString();
                        //}

                        //string sM2TimeT = string.Empty;
                        //if (dtMerge.Rows[i]["M2_TIME_T"].ToString().Length > 5)
                        //{
                        //    sM2TimeT = dtMerge.Rows[i]["M2_TIME_T"].ToString().Substring(11, 5);
                        //}
                        //else if (dtMerge.Rows[i]["M2_TIME_T"].ToString().Length == 5)
                        //{
                        //    sM2TimeT = dtMerge.Rows[i]["M2_TIME_T"].ToString();
                        //}
                        #endregion

                        string sM2TimeF = dtMerge.Rows[i]["M2_TIME_F"].ToString();
                        if (DateTime.TryParse(sM2TimeF, out DateTime dtM2TimeF))
                            sM2TimeF = dtM2TimeF.ToString("HH:mm");
                        else
                            sM2TimeF = "";

                        string sM2TimeT = dtMerge.Rows[i]["M2_TIME_T"].ToString();
                        if (DateTime.TryParse(sM2TimeT, out DateTime dtM2TimeT))
                            sM2TimeT = dtM2TimeT.ToString("HH:mm");
                        else
                            sM2TimeT = "";

                        string sM2User = dtMerge.Rows[i]["M2_USER"].ToString();
                        string sM2UserId = dtMerge.Rows[i]["M2_USER_ID"].ToString();
                        double dM2Charge = String.IsNullOrEmpty(dtMerge.Rows[i]["M2_CHARGE"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M2_CHARGE"]);
                        double dM2DPut = String.IsNullOrEmpty(dtMerge.Rows[i]["M2_D_PUT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M2_D_PUT"]);
                        double dM2DIn = String.IsNullOrEmpty(dtMerge.Rows[i]["M2_D_IN"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M2_D_IN"]);
                        double dM2CpMove = String.IsNullOrEmpty(dtMerge.Rows[i]["M2_CP_MOVE"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M2_CP_MOVE"]);
                        double dM2CpPut = String.IsNullOrEmpty(dtMerge.Rows[i]["M2_CP_PUT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M2_CP_PUT"]);
                        double dM2CpStock = String.IsNullOrEmpty(dtMerge.Rows[i]["M2_CP_STOCK"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M2_CP_STOCK"]);
                        string sEntId = FmMainToolBar2.UserID;
                        string sMfyId = FmMainToolBar2.UserID;

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        #region MARIADB
                        //strSql.AppendLine(" INSERT INTO MAKE_2 ");
                        //strSql.AppendLine("           ( MAKENO ");
                        //strSql.AppendLine("           , MAKENO_LM ");
                        //strSql.AppendLine("           , M2_TIME_F ");
                        //strSql.AppendLine("           , M2_TIME_T ");
                        //strSql.AppendLine("           , M2_USER ");
                        //strSql.AppendLine("           , M2_USER_ID ");
                        //strSql.AppendLine("           , M2_CHARGE ");
                        //strSql.AppendLine("           , M2_D_PUT ");
                        //strSql.AppendLine("           , M2_D_IN ");
                        //strSql.AppendLine("           , M2_CP_MOVE ");
                        //strSql.AppendLine("           , M2_CP_PUT ");
                        //strSql.AppendLine("           , M2_CP_STOCK ");
                        //strSql.AppendLine("           , ENT_DT ");
                        //strSql.AppendLine("           , ENT_ID ");
                        //strSql.AppendLine("           , MFY_DT ");
                        //strSql.AppendLine("           , MFY_ID ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("      VALUES ");
                        //strSql.AppendLine("           ( " + sMakeNo + " ");
                        //strSql.AppendLine("           , " + dMakeNoLm + "");
                        //strSql.AppendLine("           , '" + sM2TimeF + "' ");
                        //strSql.AppendLine("           , '" + sM2TimeT + "' ");
                        //strSql.AppendLine("           , '" + sM2User + "' ");
                        //strSql.AppendLine("           , '" + sM2UserId + "' ");
                        //strSql.AppendLine("           , " + dM2Charge + " ");
                        //strSql.AppendLine("           , " + dM2DPut + " ");
                        //strSql.AppendLine("           , " + dM2DIn + " ");
                        //strSql.AppendLine("           , " + dM2CpMove + " ");
                        //strSql.AppendLine("           , " + dM2CpPut + " ");
                        //strSql.AppendLine("           , " + dM2CpStock + " ");
                        //strSql.AppendLine("           , CONVERT([varchar](20),getdate(),(21)) ");
                        //strSql.AppendLine("           , '" + sEntId + "' ");
                        //strSql.AppendLine("           , CONVERT([varchar](20),getdate(),(21)) ");
                        //strSql.AppendLine("           , '" + sMfyId + "') ");
                        //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                        //strSql.AppendLine("             M2_TIME_F = '" + sM2TimeF + "' ");
                        //strSql.AppendLine("           , M2_TIME_T = '" + sM2TimeT + "' ");
                        //strSql.AppendLine("           , M2_USER = '" + sM2User + "' ");
                        //strSql.AppendLine("           , M2_USER_ID = '" + sM2UserId + "' ");
                        //strSql.AppendLine("           , M2_CHARGE = " + dM2Charge + " ");
                        //strSql.AppendLine("           , M2_D_PUT = " + dM2DPut + " ");
                        //strSql.AppendLine("           , M2_D_IN = " + dM2DIn + " ");
                        //strSql.AppendLine("           , M2_CP_MOVE = " + dM2CpMove + " ");
                        //strSql.AppendLine("           , M2_CP_PUT = " + dM2CpPut + " ");
                        //strSql.AppendLine("           , M2_CP_STOCK = " + dM2CpStock + " ");
                        //strSql.AppendLine("           , MFY_DT = NOW() ");
                        //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                        #endregion


                        strSql.AppendLine("IF EXISTS(SELECT * FROM MAKE_2 WHERE MAKENO = " + sMakeNo + "   AND MAKENO_LM = " + dMakeNoLm + ")");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         UPDATE MAKE_2");
                        strSql.AppendLine("            SET M2_TIME_F = '" + sM2TimeF + "'");
                        strSql.AppendLine("              , M2_TIME_T = '" + sM2TimeT + "'");
                        strSql.AppendLine("              , M2_USER = '" + sM2User + "'");
                        strSql.AppendLine("              , M2_USER_ID = '" + sM2UserId + "'");
                        strSql.AppendLine("              , M2_CHARGE = " + dM2Charge + "");
                        strSql.AppendLine("              , M2_D_PUT = " + dM2DPut + "");
                        strSql.AppendLine("              , M2_D_IN = " + dM2DIn + "");
                        strSql.AppendLine("              , M2_CP_MOVE = " + dM2CpMove + "");
                        strSql.AppendLine("              , M2_CP_PUT = " + dM2CpPut + "");
                        strSql.AppendLine("              , M2_CP_STOCK = " + dM2CpStock + "");
                        strSql.AppendLine("              , MFY_DT = CONVERT([varchar](20),getdate(),(21))");
                        strSql.AppendLine("              , MFY_ID = '" + sMfyId + "'");
                        strSql.AppendLine("          WHERE MAKENO = " + sMakeNo + "   AND MAKENO_LM = " + dMakeNoLm + "");
                        strSql.AppendLine("   END");
                        strSql.AppendLine("ELSE");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         INSERT INTO MAKE_2");
                        strSql.AppendLine("              (MAKENO");
                        strSql.AppendLine("              , MAKENO_LM");
                        strSql.AppendLine("              , M2_TIME_F");
                        strSql.AppendLine("              , M2_TIME_T");
                        strSql.AppendLine("              , M2_USER");
                        strSql.AppendLine("              , M2_USER_ID");
                        strSql.AppendLine("              , M2_CHARGE");
                        strSql.AppendLine("              , M2_D_PUT");
                        strSql.AppendLine("              , M2_D_IN");
                        strSql.AppendLine("              , M2_CP_MOVE");
                        strSql.AppendLine("              , M2_CP_PUT");
                        strSql.AppendLine("              , M2_CP_STOCK");
                        strSql.AppendLine("              , ENT_DT");
                        strSql.AppendLine("              , ENT_ID");
                        strSql.AppendLine("              , MFY_DT");
                        strSql.AppendLine("              , MFY_ID");
                        strSql.AppendLine("              )");
                        strSql.AppendLine("         VALUES");
                        strSql.AppendLine("              (" + sMakeNo + "");
                        strSql.AppendLine("              , " + dMakeNoLm + "");
                        strSql.AppendLine("              , '" + sM2TimeF + "'");
                        strSql.AppendLine("              , '" + sM2TimeT + "'");
                        strSql.AppendLine("              , '" + sM2User + "'");
                        strSql.AppendLine("              , '" + sM2UserId + "'");
                        strSql.AppendLine("              , " + dM2Charge + "");
                        strSql.AppendLine("              , " + dM2DPut + "");
                        strSql.AppendLine("              , " + dM2DIn + "");
                        strSql.AppendLine("              , " + dM2CpMove + "");
                        strSql.AppendLine("              , " + dM2CpPut + "");
                        strSql.AppendLine("              , " + dM2CpStock + "");
                        strSql.AppendLine("              , CONVERT([varchar](20), getdate(), (21))");
                        strSql.AppendLine("              , '" + sEntId + "'");
                        strSql.AppendLine("              , CONVERT([varchar](20),getdate(),(21))");
                        strSql.AppendLine("              , '" + sMfyId + "')");
                        strSql.AppendLine("   END");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();


                        #region[재고이동]
                        if (dM2CpMove > 0)
                        {
                            strSql.Clear();
                            strSql.AppendLine(" ");

                            strSql.AppendLine("         DELETE FROM MESURING ");
                            strSql.AppendLine("         WHERE J_Date = '" + ymd + "' ");
                            strSql.AppendLine("         AND   Sun = 99 ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();


                            strSql.Clear();
                            strSql.AppendLine(" ");

                            strSql.AppendLine("         INSERT INTO MESURING");
                            strSql.AppendLine("              ( KeraType");
                            strSql.AppendLine("              , j_check ");
                            strSql.AppendLine("              , MaipCherID");
                            strSql.AppendLine("              , J_AssignID");
                            strSql.AppendLine("              , J_Company");
                            strSql.AppendLine("              , Sun");
                            strSql.AppendLine("              , J_Date");
                            strSql.AppendLine("              , FirstTime");
                            strSql.AppendLine("              , SecondTime");
                            strSql.AppendLine("              , FirstWeight");
                            strSql.AppendLine("              , SecondWeight");
                            strSql.AppendLine("              , Weight");
                            strSql.AppendLine("              , J_Serial");
                            strSql.AppendLine("              , Gubun1");
                            strSql.AppendLine("              , J_BNum");
                            strSql.AppendLine("              , K_Name");
                            strSql.AppendLine("              , iChaGam, OChaGam,iGamga,OGamga,iWeight,Oweight,iDanga,ODanga,iKongkep,Okongkep");
                            strSql.AppendLine("              , j_state2,regionstart,regiondest");
                            strSql.AppendLine("              , transportdanga,transportkumak,transportperson,transportcustom,transportc_serial,transportbigo,customweight,lossweight,transportjungsan");
                            strSql.AppendLine("              , magam_flag,ipchulgo_maipid,ipchulgo_machulid,weight_gubun,lengthsid,blno,containerno,damage");
                            strSql.AppendLine("              ,unit_prc_chg_yn, seak_poham, usercode, buseocode, p_id, j_garage, j_id, kyeryang12, driver_inout, agree_date, gumsu_serial, halinyul, suryang");
                            strSql.AppendLine("              )");
                            strSql.AppendLine("         VALUES");
                            strSql.AppendLine("              ('출고'");
                            strSql.AppendLine("              ,' ' ");
                            strSql.AppendLine("              , 0");
                            strSql.AppendLine("              , 6303870006");
                            strSql.AppendLine("              , '재고이동'");
                            strSql.AppendLine("              , 99");
                            strSql.AppendLine("              , '" + ymd + "' ");
                            strSql.AppendLine("              , '" + Firsthhmmss + "' ");
                            strSql.AppendLine("              , '" + Firsthhmmss + "' ");
                            strSql.AppendLine("              , 1000");
                            strSql.AppendLine("              , " + dM2CpMove + " * 500 + 1000 ");
                            strSql.AppendLine("              , " + dM2CpMove + " * 500 ");
                            strSql.AppendLine("              , 5059072");
                            strSql.AppendLine("              , '슈-고철B'");
                            strSql.AppendLine("              , '1111'");
                            strSql.AppendLine("              , '상품'");
                            strSql.AppendLine("              , 0,0,0, 0,0, (" + dM2CpMove + " * 500) ,0,0,0,0");
                            strSql.AppendLine("              , ' ',' ',' '");
                            strSql.AppendLine("              , 0,0,' ',' ',0,' ',(" + dM2CpMove + " * 500),0,0,0,0,0,0,0,' ',' ',0 ");
                            strSql.AppendLine("              ,'N','N',2001,0,100,1,0,1,0,'0000-01-01',1051,0,0 ");
                            strSql.AppendLine("              )");
                           
                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                        }
                        #endregion[재고이동]

                    }
                    Cursor = Cursors.Default;
                }

                #endregion[길로틴 저장부분]

                #region[슈레더 저장부분]

                if (GridViewShreder.RowCount > 0)
                {
                    Cursor = Cursors.WaitCursor;

                    DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridShreder.DataSource);
                    DataTable dtMerge = dsSave.Tables[0];

                    for (int i = 0; i < dtMerge.Rows.Count; i++)
                    {
                        if (i == 0 || i == dtMerge.Rows.Count - 1)
                        {
                            string chksM3TimeF = dtMerge.Rows[i]["M3_TIME_F"]?.ToString();
                            string chksM3TimeT = dtMerge.Rows[i]["M3_TIME_F"]?.ToString();
                            string chksM3In = dtMerge.Rows[i]["M3_IN"].ToString();
                            string chksM3Cp = dtMerge.Rows[i]["M3_CP"].ToString();
                            string chksM3Etc = dtMerge.Rows[i]["M3_Dmt"].ToString();
                            string chksM3WGT = dtMerge.Rows[i]["M3_WGT"].ToString();
                            string chksM3BAN = dtMerge.Rows[i]["M3_BAN"].ToString();
                            string chksM3LIGHT = dtMerge.Rows[i]["M3_LIGHT"].ToString();
                            string chksM3CarAvg = dtMerge.Rows[i]["M3_CAR_AVG"].ToString();
                            string chksM3Line = dtMerge.Rows[i]["M3_LINE"].ToString();
                            string chksM3Put = dtMerge.Rows[i]["M3_PUT"].ToString();
                            string chksM3Op = dtMerge.Rows[i]["M3_OP"].ToString();
                            string chksM3M3Fine = dtMerge.Rows[i]["M3_M3_FIND"].ToString();

                            if (string.IsNullOrEmpty(chksM3TimeF) & string.IsNullOrEmpty(chksM3TimeT)
                              & string.IsNullOrEmpty(chksM3In) & string.IsNullOrEmpty(chksM3Cp)
                              & string.IsNullOrEmpty(chksM3Etc) & string.IsNullOrEmpty(chksM3WGT)
                              & string.IsNullOrEmpty(chksM3BAN) & string.IsNullOrEmpty(chksM3LIGHT)
                              & string.IsNullOrEmpty(chksM3CarAvg) & string.IsNullOrEmpty(chksM3Line)
                              & string.IsNullOrEmpty(chksM3Op) & string.IsNullOrEmpty(chksM3M3Fine)
                              & string.IsNullOrEmpty(chksM3Put) )
                            {
                                continue;
                            }
                            
                        }

                        string sMakeNo = dMakeNo.ToString();
                        double dMakeNoLm = 0;

                        if (String.IsNullOrEmpty(dtMerge.Rows[i]["MAKENO_LM"].ToString()))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" SELECT MAX(A.MAKENO_LM)");
                            strSql.AppendLine("   FROM MAKE_3 A ");
                            strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();

                            if (String.IsNullOrEmpty(cmd.ExecuteScalar().ToString()))
                            {
                                dMakeNoLm = 1;
                            }
                            else
                            {
                                dMakeNoLm = Convert.ToDouble(cmd.ExecuteScalar()) + 1;
                            }
                        }
                        else
                        {
                            dMakeNoLm = Convert.ToDouble(dtMerge.Rows[i]["MAKENO_LM"]);
                        }

                        #region 2021-08-31 이전
                        //string sM3TimeF = string.Empty;
                        //if (dtMerge.Rows[i]["M3_TIME_F"].ToString().Length > 5)
                        //{
                        //    sM3TimeF = dtMerge.Rows[i]["M3_TIME_F"].ToString().Substring(11, 5);
                        //}
                        //else if (dtMerge.Rows[i]["M3_TIME_F"].ToString().Length == 5)
                        //{
                        //    sM3TimeF = dtMerge.Rows[i]["M3_TIME_F"].ToString();
                        //}

                        //string sM3TimeT = string.Empty;
                        //if (dtMerge.Rows[i]["M3_TIME_T"].ToString().Length > 5)
                        //{
                        //    sM3TimeT = dtMerge.Rows[i]["M3_TIME_T"].ToString().Substring(11, 5);
                        //}
                        //else if (dtMerge.Rows[i]["M3_TIME_T"].ToString().Length == 5)
                        //{
                        //    sM3TimeT = dtMerge.Rows[i]["M3_TIME_T"].ToString();
                        //}

                        //string sM3Time = string.Empty;
                        //if (dtMerge.Rows[i]["M3_TIME"].ToString().Length > 5)
                        //{
                        //    sM3Time = dtMerge.Rows[i]["M3_TIME"].ToString().Substring(11, 5);
                        //}
                        //else if (dtMerge.Rows[i]["M3_TIME"].ToString().Length == 5)
                        //{
                        //    sM3Time = dtMerge.Rows[i]["M3_TIME"].ToString();
                        //}
                        #endregion

                        string sM3TimeF = dtMerge.Rows[i]["M3_TIME_F"].ToString();
                        if (DateTime.TryParse(sM3TimeF, out DateTime dtM3TimeF))
                            sM3TimeF = dtM3TimeF.ToString("HH:mm");
                        else
                            sM3TimeF = "";

                        string sM3TimeT = dtMerge.Rows[i]["M3_TIME_T"].ToString();
                        if (DateTime.TryParse(sM3TimeT, out DateTime dtM3TimeT))
                            sM3TimeT = dtM3TimeT.ToString("HH:mm");
                        else
                            sM3TimeT = "";

                        string sM3Time = dtMerge.Rows[i]["M3_TIME"].ToString();
                        if (DateTime.TryParse(sM3Time, out DateTime dtM3Time))
                            sM3Time = dtM3Time.ToString("HH:mm");
                        else
                            sM3Time = "";

                        //string sM3Dmt = dtMerge.Rows[i]["M3_Dmt"].ToString();
                        //if (DateTime.TryParse(sM3Dmt, out DateTime dtM3TimeEtc))
                        //    sM3Dmt = dtM3TimeEtc.ToString("HH:mm");
                        //else
                        //    sM3Dmt = "";

                        double dM3In = String.IsNullOrEmpty(dtMerge.Rows[i]["M3_IN"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M3_IN"]);
                        double dM3Cp = String.IsNullOrEmpty(dtMerge.Rows[i]["M3_CP"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M3_CP"]);
                        double sM3Dmt = String.IsNullOrEmpty(dtMerge.Rows[i]["M3_Dmt"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M3_Dmt"]);
                        double dM3WGT = String.IsNullOrEmpty(dtMerge.Rows[i]["M3_WGT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M3_WGT"]);
                        double dM3BAN = String.IsNullOrEmpty(dtMerge.Rows[i]["M3_BAN"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M3_BAN"]);
                        double dM3LIGHT = String.IsNullOrEmpty(dtMerge.Rows[i]["M3_LIGHT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M3_LIGHT"]);
                        double dM3CarAvg = String.IsNullOrEmpty(dtMerge.Rows[i]["M3_CAR_AVG"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M3_CAR_AVG"]);
                        string sM3Line = dtMerge.Rows[i]["M3_LINE"]?.ToString();
                        string sM3Put = dtMerge.Rows[i]["M3_PUT"]?.ToString();
                        string sM3Op = dtMerge.Rows[i]["M3_OP"]?.ToString();
                        string sM3M3Fine = dtMerge.Rows[i]["M3_M3_FIND"]?.ToString();

                        string sEntId = FmMainToolBar2.UserID;
                        string sMfyId = FmMainToolBar2.UserID;

                        strSql.Clear();
                        strSql.AppendLine(" ");

                        #region mariaDB
                        //strSql.AppendLine(" INSERT INTO MAKE_3 ");
                        //strSql.AppendLine("           ( MAKENO ");
                        //strSql.AppendLine("           , MAKENO_LM ");
                        //strSql.AppendLine("           , M3_TIME_F ");
                        //strSql.AppendLine("           , M3_TIME_T ");
                        //strSql.AppendLine("           , M3_TIME ");
                        //strSql.AppendLine("           , M3_IN ");
                        //strSql.AppendLine("           , M3_CP");
                        //strSql.AppendLine("           , M3_Dmt");
                        //strSql.AppendLine("           , M3_WGT");
                        //strSql.AppendLine("           , M3_CAR_AVG");
                        //strSql.AppendLine("           , M3_LINE");
                        //strSql.AppendLine("           , M3_PUT");
                        //strSql.AppendLine("           , M3_OP");
                        //strSql.AppendLine("           , M3_M3_FIND");
                        //strSql.AppendLine("           , ENT_DT ");
                        //strSql.AppendLine("           , ENT_ID ");
                        //strSql.AppendLine("           , MFY_DT ");
                        //strSql.AppendLine("           , MFY_ID ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("      VALUES ");
                        //strSql.AppendLine("           ( " + sMakeNo + " ");
                        //strSql.AppendLine("           , " + dMakeNoLm + "");
                        //strSql.AppendLine("           , '" + sM3TimeF + "' ");
                        //strSql.AppendLine("           , '" + sM3TimeT + "' ");
                        //strSql.AppendLine("           , '" + sM3Time + "' ");
                        //strSql.AppendLine("           , " + dM3In + "");
                        //strSql.AppendLine("           , " + dM3Cp + "");
                        //strSql.AppendLine("           , " + dM3Etc + "");
                        //strSql.AppendLine("           , " + dM3WGT + "");
                        //strSql.AppendLine("           , " + dM3CarAvg + "");
                        //strSql.AppendLine("           , '" + sM3Line + "' ");
                        //strSql.AppendLine("           , '" + sM3Put + "' ");
                        //strSql.AppendLine("           , '" + sM3Op + "' ");
                        //strSql.AppendLine("           , '" + sM3M3Fine + "' ");
                        //strSql.AppendLine("           , CONVERT([varchar](20), getdate(), (21))");
                        //strSql.AppendLine("           , '" + sEntId + "' ");
                        //strSql.AppendLine("           , CONVERT([varchar](20), getdate(), (21))");
                        //strSql.AppendLine("           , '" + sMfyId + "') ");
                        //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                        //strSql.AppendLine("             M3_TIME_F = '" + sM3TimeF + "' ");
                        //strSql.AppendLine("           , M3_TIME_T = '" + sM3TimeT + "' ");
                        //strSql.AppendLine("           , M3_TIME = '" + sM3Time + "' ");
                        //strSql.AppendLine("           , M3_IN = " + dM3In + " ");
                        //strSql.AppendLine("           , M3_CP = " + dM3Cp + " ");
                        //strSql.AppendLine("           , M3_Dmt = " + dM3Etc + " ");
                        //strSql.AppendLine("           , M3_WGT = " + dM3WGT + " ");
                        //strSql.AppendLine("           , M3_CAR_AVG = " + dM3CarAvg + " ");
                        //strSql.AppendLine("           , M3_LINE = '" + sM3Line + "' ");
                        //strSql.AppendLine("           , M3_PUT = '" + sM3Put + "' ");
                        //strSql.AppendLine("           , M3_OP = '" + sM3Op + "' ");
                        //strSql.AppendLine("           , M3_M3_FIND = '" + sM3M3Fine + "' ");
                        //strSql.AppendLine("           , MFY_DT = NOW() ");
                        //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                        #endregion

                        strSql.AppendLine("IF EXISTS(SELECT * FROM MAKE_3 WHERE MAKENO = " + sMakeNo + "   AND MAKENO_LM = " + dMakeNoLm + ")");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         UPDATE MAKE_3");
                        strSql.AppendLine("            SET M3_TIME_F = '" + sM3TimeF + "' ");
                        strSql.AppendLine("              , M3_TIME_T = '" + sM3TimeT + "' ");
                        strSql.AppendLine("              , M3_TIME = '" + sM3Time + "' ");
                        strSql.AppendLine("              , M3_IN = " + dM3In + " ");
                        strSql.AppendLine("              , M3_CP = " + dM3Cp + " ");
                        strSql.AppendLine("              , M3_Dmt = '" + sM3Dmt + "' ");
                        strSql.AppendLine("              , M3_WGT = " + dM3WGT + " ");
                        strSql.AppendLine("              , M3_BAN = " + dM3BAN + " ");
                        strSql.AppendLine("              , M3_LIGHT = " + dM3LIGHT + " ");
                        strSql.AppendLine("              , M3_CAR_AVG = " + dM3CarAvg + " ");
                        strSql.AppendLine("              , M3_LINE = '" + sM3Line + "' ");
                        strSql.AppendLine("              , M3_PUT = '" + sM3Put + "' ");
                        strSql.AppendLine("              , M3_OP = '" + sM3Op + "' ");
                        strSql.AppendLine("              , M3_M3_FIND = '" + sM3M3Fine + "' ");
                        strSql.AppendLine("              , MFY_DT = CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("              , MFY_ID = '" + sMfyId + "' ");
                        strSql.AppendLine("          WHERE MAKENO = " + sMakeNo + "   AND MAKENO_LM = " + dMakeNoLm);
                        strSql.AppendLine("   END");
                        strSql.AppendLine("ELSE");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         INSERT INTO MAKE_3 ");
                        strSql.AppendLine("                   ( MAKENO ");
                        strSql.AppendLine("                   , MAKENO_LM ");
                        strSql.AppendLine("                   , M3_TIME_F ");
                        strSql.AppendLine("                   , M3_TIME_T ");
                        strSql.AppendLine("                   , M3_TIME ");
                        strSql.AppendLine("                   , M3_IN ");
                        strSql.AppendLine("                   , M3_CP");
                        strSql.AppendLine("                   , M3_Dmt");
                        strSql.AppendLine("                   , M3_WGT");
                        strSql.AppendLine("                   , M3_BAN");
                        strSql.AppendLine("                   , M3_LIGHT");
                        strSql.AppendLine("                   , M3_CAR_AVG");
                        strSql.AppendLine("                   , M3_LINE");
                        strSql.AppendLine("                   , M3_PUT");
                        strSql.AppendLine("                   , M3_OP");
                        strSql.AppendLine("                   , M3_M3_FIND");
                        strSql.AppendLine("                   , ENT_DT ");
                        strSql.AppendLine("                   , ENT_ID ");
                        strSql.AppendLine("                   , MFY_DT ");
                        strSql.AppendLine("                   , MFY_ID ");
                        strSql.AppendLine("                   ) ");
                        strSql.AppendLine("              VALUES ");
                        strSql.AppendLine("                   ( " + sMakeNo + " ");
                        strSql.AppendLine("                   , " + dMakeNoLm + "");
                        strSql.AppendLine("                   , '" + sM3TimeF + "' ");
                        strSql.AppendLine("                   , '" + sM3TimeT + "' ");
                        strSql.AppendLine("                   , '" + sM3Time + "' ");
                        strSql.AppendLine("                   , " + dM3In + "");
                        strSql.AppendLine("                   , " + dM3Cp + "");
                        strSql.AppendLine("                   , '" + sM3Dmt + "'");
                        strSql.AppendLine("                   , " + dM3WGT + "");
                        strSql.AppendLine("                   , " + dM3BAN + "");
                        strSql.AppendLine("                   , " + dM3LIGHT + "");
                        strSql.AppendLine("                   , " + dM3CarAvg + "");
                        strSql.AppendLine("                   , '" + sM3Line + "' ");
                        strSql.AppendLine("                   , '" + sM3Put + "' ");
                        strSql.AppendLine("                   , '" + sM3Op + "' ");
                        strSql.AppendLine("                   , '" + sM3M3Fine + "' ");
                        strSql.AppendLine("                   , CONVERT([varchar](20), getdate(), (21))");
                        strSql.AppendLine("                   , '" + sEntId + "' ");
                        strSql.AppendLine("                   , CONVERT([varchar](20), getdate(), (21))");
                        strSql.AppendLine("                   , '" + sMfyId + "') ");
                        strSql.AppendLine("   END");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }
                    Cursor = Cursors.Default;
                }

                #endregion[슈레더 저장부분]

                #region[정비 저장부분]
                
                if (GridViewMaintenance.RowCount > 0)
                {
                    Cursor = Cursors.WaitCursor;

                    DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridMaintenance.DataSource);
                    DataTable dtMerge = dsSave.Tables[0];

                    for (int i = 0; i < dtMerge.Rows.Count; i++)
                    {
                        if (i == 0 || i == dtMerge.Rows.Count - 1)
                        {
                            string chksM5TimeF = dtMerge.Rows[i]["M5_TIME_F"]?.ToString();
                            string chksM5TimeT = dtMerge.Rows[i]["M5_TIME_F"]?.ToString();
                            string chksM5OverTime = dtMerge.Rows[i]["M5_OVERTIME"].ToString();
                            string chksM5DeviceCd = dtMerge.Rows[i]["M5_DEVICE_CD"].ToString();
                            string chksM5Device = dtMerge.Rows[i]["M5_DEVICE"].ToString();
                            string chksM5ChrgId = dtMerge.Rows[i]["M5_CHRG_ID"].ToString();
                            string chksM5ChrgNm = dtMerge.Rows[i]["M5_CHRG_NM"].ToString();

                            if (string.IsNullOrEmpty(chksM5TimeF) & string.IsNullOrEmpty(chksM5TimeT)
                              & string.IsNullOrEmpty(chksM5OverTime) & string.IsNullOrEmpty(chksM5Device)
                              & string.IsNullOrEmpty(chksM5DeviceCd) & string.IsNullOrEmpty(chksM5ChrgId)
                              & string.IsNullOrEmpty(chksM5ChrgId) & string.IsNullOrEmpty(chksM5ChrgNm))
                            {
                                continue;
                            }
                        }

                        string sMakeNo = dMakeNo.ToString();
                        double dMakeNoLm = 0;

                        if (String.IsNullOrEmpty(dtMerge.Rows[i]["MAKENO_LM"].ToString()))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" SELECT MAX(A.MAKENO_LM)");
                            strSql.AppendLine("   FROM MAKE_5 A ");
                            strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();

                            if (String.IsNullOrEmpty(cmd.ExecuteScalar().ToString()))
                            {
                                dMakeNoLm = 1;
                            }
                            else
                            {
                                dMakeNoLm = Convert.ToDouble(cmd.ExecuteScalar()) + 1;
                            }
                        }
                        else
                        {
                            dMakeNoLm = Convert.ToDouble(dtMerge.Rows[i]["MAKENO_LM"]);
                        }

                        #region 2021-08-31 이전코드
                        //string sM5TimeF = string.Empty;
                        //if (dtMerge.Rows[i]["M5_TIME_F"].ToString().Length > 5)
                        //{
                        //    sM5TimeF = dtMerge.Rows[i]["M5_TIME_F"].ToString().Substring(11, 5);
                        //}
                        //else if (dtMerge.Rows[i]["M5_TIME_F"].ToString().Length == 5)
                        //{
                        //    sM5TimeF = dtMerge.Rows[i]["M5_TIME_F"].ToString();
                        //}

                        //string sM5TimeT = string.Empty;
                        //if (dtMerge.Rows[i]["M5_TIME_T"].ToString().Length > 5)
                        //{
                        //    sM5TimeT = dtMerge.Rows[i]["M5_TIME_T"].ToString().Substring(11, 5);
                        //}
                        //else if (dtMerge.Rows[i]["M5_TIME_T"].ToString().Length == 5)
                        //{
                        //    sM5TimeT = dtMerge.Rows[i]["M5_TIME_T"].ToString();
                        //}
                        #endregion

                        string sM5TimeF = dtMerge.Rows[i]["M5_TIME_F"].ToString();
                        if (DateTime.TryParse(sM5TimeF, out DateTime dtM5TimeF))
                            sM5TimeF = dtM5TimeF.ToString("HH:mm");
                        else
                            sM5TimeF = "";

                        string sM5TimeT = dtMerge.Rows[i]["M5_TIME_T"].ToString();
                        if (DateTime.TryParse(sM5TimeT, out DateTime dtM5TimeT))
                            sM5TimeT = dtM5TimeT.ToString("HH:mm");
                        else
                            sM5TimeT = "";

                        string sM5OverTime = dtMerge.Rows[i]["M5_OVERTIME"].ToString();
                        //string sM5Device = dtMerge.Rows[i]["M5_DEVICE"].ToString();
                        string sM5DeviceCd = dtMerge.Rows[i]["M5_DEVICE_CD"].ToString();
                        string sDEVICE = string.Empty;
                        if (string.IsNullOrEmpty(sM5DeviceCd))
                        {
                            sDEVICE = dtMerge.Rows[i]["M5_DEVICE"].ToString();
                        }
                        else
                        {
                            DataTable dtNM = SetM5Device(sM5DeviceCd);
                            sDEVICE = dtNM.Rows[0]["EQUIP_NM"]?.ToString();
                        }
                        string sM5Device = sDEVICE;
                        string sM5ChrgId = dtMerge.Rows[i]["M5_CHRG_ID"].ToString();
                        //string sM5ChrgNm = dtMerge.Rows[i]["M5_CHRG_NM"].ToString();
                        string sCHRG = string.Empty;
                        if (string.IsNullOrEmpty(sM5ChrgId))
                        {
                            sCHRG = dtMerge.Rows[i]["M5_CHRG_NM"].ToString();
                        }
                        else
                        {
                            DataTable dtCH = SetM5CHRG(sM5ChrgId);
                            sCHRG = dtCH.Rows[0]["EMP_NM"]?.ToString();
                        }
                        string sM5ChrgNm = sCHRG;
                        string sEntId = FmMainToolBar2.UserID;
                        string sMfyId = FmMainToolBar2.UserID;

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        #region MARIADB
                        //strSql.AppendLine(" INSERT INTO MAKE_5 ");
                        //strSql.AppendLine("           ( MAKENO ");
                        //strSql.AppendLine("           , MAKENO_LM ");
                        //strSql.AppendLine("           , M5_TIME_F ");
                        //strSql.AppendLine("           , M5_TIME_T ");
                        //strSql.AppendLine("           , M5_OVERTIME ");
                        //strSql.AppendLine("           , M5_DEVICE ");
                        //strSql.AppendLine("           , M5_DEVICE_CD ");
                        //strSql.AppendLine("           , M5_CHRG_ID ");
                        //strSql.AppendLine("           , M5_CHRG_NM ");
                        //strSql.AppendLine("           , ENT_DT ");
                        //strSql.AppendLine("           , ENT_ID ");
                        //strSql.AppendLine("           , MFY_DT ");
                        //strSql.AppendLine("           , MFY_ID ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("      VALUES ");
                        //strSql.AppendLine("           ( " + sMakeNo + " ");
                        //strSql.AppendLine("           , " + dMakeNoLm + "");
                        //strSql.AppendLine("           , '" + sM5TimeF + "' ");
                        //strSql.AppendLine("           , '" + sM5TimeT + "' ");
                        //strSql.AppendLine("           , '" + sM5OverTime + "' ");
                        //strSql.AppendLine("           , '" + sM5Device + "' ");
                        //strSql.AppendLine("           , '" + sM5DeviceCd + "' ");
                        //strSql.AppendLine("           , '" + sM5ChrgId + "' ");
                        //strSql.AppendLine("           , '" + sM5ChrgNm + "' ");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sEntId + "' ");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sMfyId + "') ");
                        //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                        //strSql.AppendLine("             M5_TIME_F = '" + sM5TimeF + "' ");
                        //strSql.AppendLine("           , M5_TIME_T = '" + sM5TimeT + "' ");
                        //strSql.AppendLine("           , M5_OVERTIME = '" + sM5OverTime + "' ");
                        //strSql.AppendLine("           , M5_DEVICE = '" + sM5Device + "' ");
                        //strSql.AppendLine("           , M5_DEVICE_CD = '" + sM5DeviceCd + "' ");
                        //strSql.AppendLine("           , M5_CHRG_ID = '" + sM5ChrgId + "' ");
                        //strSql.AppendLine("           , M5_CHRG_NM = '" + sM5ChrgNm + "' ");
                        //strSql.AppendLine("           , MFY_DT = NOW() ");
                        //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                        #endregion

                        strSql.AppendLine("IF EXISTS(SELECT * FROM MAKE_5 WHERE MAKENO = " + sMakeNo + "   AND MAKENO_LM = " + dMakeNoLm + ")");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         UPDATE MAKE_5");
                        strSql.AppendLine("            SET M5_TIME_F = '" + sM5TimeF + "' ");
                        strSql.AppendLine("              , M5_TIME_T = '" + sM5TimeT + "' ");
                        strSql.AppendLine("              , M5_OVERTIME = '" + sM5OverTime + "' ");
                        strSql.AppendLine("              , M5_DEVICE = '" + sM5Device + "' ");
                        strSql.AppendLine("              , M5_DEVICE_CD = '" + sM5DeviceCd + "' ");
                        strSql.AppendLine("              , M5_CHRG_ID = '" + sM5ChrgId + "' ");
                        strSql.AppendLine("              , M5_CHRG_NM = '" + sM5ChrgNm + "' ");
                        strSql.AppendLine("              , MFY_DT = CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("              , MFY_ID = '" + sMfyId + "' ");
                        strSql.AppendLine("          WHERE MAKENO = " + sMakeNo + "   AND MAKENO_LM = " + dMakeNoLm);
                        strSql.AppendLine("   END");
                        strSql.AppendLine("ELSE");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         INSERT INTO MAKE_5 ");
                        strSql.AppendLine("                   ( MAKENO ");
                        strSql.AppendLine("                   , MAKENO_LM ");
                        strSql.AppendLine("                   , M5_TIME_F ");
                        strSql.AppendLine("                   , M5_TIME_T ");
                        strSql.AppendLine("                   , M5_OVERTIME ");
                        strSql.AppendLine("                   , M5_DEVICE ");
                        strSql.AppendLine("                   , M5_DEVICE_CD ");
                        strSql.AppendLine("                   , M5_CHRG_ID ");
                        strSql.AppendLine("                   , M5_CHRG_NM ");
                        strSql.AppendLine("                   , ENT_DT ");
                        strSql.AppendLine("                   , ENT_ID ");
                        strSql.AppendLine("                   , MFY_DT ");
                        strSql.AppendLine("                   , MFY_ID ");
                        strSql.AppendLine("                   ) ");
                        strSql.AppendLine("              VALUES ");
                        strSql.AppendLine("                   ( " + sMakeNo + " ");
                        strSql.AppendLine("                   , " + dMakeNoLm + "");
                        strSql.AppendLine("                   , '" + sM5TimeF + "' ");
                        strSql.AppendLine("                   , '" + sM5TimeT + "' ");
                        strSql.AppendLine("                   , '" + sM5OverTime + "' ");
                        strSql.AppendLine("                   , '" + sM5Device + "' ");
                        strSql.AppendLine("                   , '" + sM5DeviceCd + "' ");
                        strSql.AppendLine("                   , '" + sM5ChrgId + "' ");
                        strSql.AppendLine("                   , '" + sM5ChrgNm + "' ");
                        strSql.AppendLine("                   , CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("                   , '" + sEntId + "' ");
                        strSql.AppendLine("                   , CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("                   , '" + sMfyId + "') ");
                        strSql.AppendLine("   END");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }
                    Cursor = Cursors.Default;
                }

                #endregion[정비 저장부분]

                #region[비용 저장부분]

                if (GridViewCost.RowCount > 0)
                {
                    Cursor = Cursors.WaitCursor;

                    DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridCost.DataSource);
                    DataTable dtMerge = dsSave.Tables[0];

                    for (int i = 0; i < dtMerge.Rows.Count; i++)
                    {
                        string sMakeNo = dMakeNo.ToString();
                        double dMakeNoLm = 0;

                        if (String.IsNullOrEmpty(dtMerge.Rows[i]["MAKENO_LM"].ToString()))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" SELECT MAX(A.MAKENO_LM)");
                            strSql.AppendLine("   FROM MAKE_6 A ");
                            strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();

                            if (String.IsNullOrEmpty(cmd.ExecuteScalar().ToString()))
                            {
                                dMakeNoLm = 1;
                            }
                            else
                            {
                                dMakeNoLm = Convert.ToDouble(cmd.ExecuteScalar()) + 1;
                            }
                        }
                        else
                        {
                            dMakeNoLm = Convert.ToDouble(dtMerge.Rows[i]["MAKENO_LM"]);
                        }

                        string sM6CvNam = dtMerge.Rows[i]["M6_CVNAM"].ToString();
                        string sM6CvNamCd = dtMerge.Rows[i]["M6_CVNAM_CD"].ToString();
                        string sM6ItNam = dtMerge.Rows[i]["M6_ITNAM"].ToString();
                        string sM6BreakReason = dtMerge.Rows[i]["M6_BREAK_REASON"].ToString();
                        double dM6Amt = String.IsNullOrEmpty(dtMerge.Rows[i]["M6_AMT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M6_AMT"].ToString());
                        string sM6Device = dtMerge.Rows[i]["M6_DEVICE"].ToString();
                        string sM6DeviceCd = dtMerge.Rows[i]["M6_DEVICE_CD"].ToString();
                        string sWDate = string.Empty;
                        if(dtMerge.Rows[i]["WDATE"].ToString().Length > 10)
                        {
                            sWDate = dtMerge.Rows[i]["WDATE"].ToString().Substring(0, 10);
                        }
                        else if(dtMerge.Rows[i]["WDATE"].ToString().Length == 10)
                        {
                            sWDate = dtMerge.Rows[i]["WDATE"].ToString();
                        }

                        string sEntId = FmMainToolBar2.UserID;
                        string sMfyId = FmMainToolBar2.UserID;

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        #region mariaDB
                        //strSql.AppendLine(" INSERT INTO MAKE_6 ");
                        //strSql.AppendLine("           ( MAKENO ");
                        //strSql.AppendLine("           , MAKENO_LM ");
                        //strSql.AppendLine("           , M6_CVNAM ");
                        //strSql.AppendLine("           , M6_CVNAM_CD ");
                        //strSql.AppendLine("           , M6_ITNAM ");
                        //strSql.AppendLine("           , M6_BREAK_REASON ");
                        //strSql.AppendLine("           , M6_AMT ");
                        //strSql.AppendLine("           , M6_DEVICE ");
                        //strSql.AppendLine("           , M6_DEVICE_CD ");
                        //strSql.AppendLine("           , WDATE ");
                        //strSql.AppendLine("           , ENT_DT ");
                        //strSql.AppendLine("           , ENT_ID ");
                        //strSql.AppendLine("           , MFY_DT ");
                        //strSql.AppendLine("           , MFY_ID ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("      VALUES ");
                        //strSql.AppendLine("           ( " + sMakeNo + " ");
                        //strSql.AppendLine("           , " + dMakeNoLm + "");
                        //strSql.AppendLine("           , '" + sM6CvNam + "' ");
                        //strSql.AppendLine("           , '" + sM6CvNamCd + "' ");
                        //strSql.AppendLine("           , '" + sM6ItNam + "' ");
                        //strSql.AppendLine("           , '" + sM6BreakReason + "' ");
                        //strSql.AppendLine("           , " + dM6Amt + " ");
                        //strSql.AppendLine("           , '" + sM6Device + "' ");
                        //strSql.AppendLine("           , '" + sM6DeviceCd + "' ");
                        //strSql.AppendLine("           , '" + sWDate + "' ");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sEntId + "' ");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sMfyId + "') ");
                        //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                        //strSql.AppendLine("             M6_CVNAM = '" + sM6CvNam + "' ");
                        //strSql.AppendLine("           , M6_CVNAM_CD = '" + sM6CvNamCd + "' ");
                        //strSql.AppendLine("           , M6_ITNAM = '" + sM6ItNam + "' ");
                        //strSql.AppendLine("           , M6_BREAK_REASON = '" + sM6BreakReason + "' ");
                        //strSql.AppendLine("           , M6_AMT = " + dM6Amt + " ");
                        //strSql.AppendLine("           , M6_DEVICE = '" + sM6Device + "' ");
                        //strSql.AppendLine("           , M6_DEVICE_CD = '" + sM6DeviceCd + "' ");
                        //strSql.AppendLine("           , WDATE = '" + sWDate + "' ");
                        //strSql.AppendLine("           , MFY_DT = NOW() ");
                        //strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");
                        #endregion

                        strSql.AppendLine("IF EXISTS(SELECT * FROM MAKE_6 WHERE MAKENO = " + sMakeNo + "   AND MAKENO_LM = " + dMakeNoLm + ")");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         UPDATE MAKE_6");
                        strSql.AppendLine("            SET M6_CVNAM = '" + sM6CvNam + "' ");
                        strSql.AppendLine("              , M6_CVNAM_CD = '" + sM6CvNamCd + "' ");
                        strSql.AppendLine("              , M6_ITNAM = '" + sM6ItNam + "' ");
                        strSql.AppendLine("              , M6_BREAK_REASON = '" + sM6BreakReason + "' ");
                        strSql.AppendLine("              , M6_AMT = " + dM6Amt + " ");
                        strSql.AppendLine("              , M6_DEVICE = '" + sM6Device + "' ");
                        strSql.AppendLine("              , M6_DEVICE_CD = '" + sM6DeviceCd + "' ");
                        strSql.AppendLine("              , WDATE = '" + sWDate + "' ");
                        strSql.AppendLine("              , MFY_DT = CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("              , MFY_ID = '" + sMfyId + "' ");
                        strSql.AppendLine("          WHERE MAKENO = " + sMakeNo + "   AND MAKENO_LM = " + dMakeNoLm);
                        strSql.AppendLine("   END");
                        strSql.AppendLine("ELSE");
                        strSql.AppendLine("   BEGIN");
                        strSql.AppendLine("         INSERT INTO MAKE_6 ");
                        strSql.AppendLine("                   ( MAKENO ");
                        strSql.AppendLine("                   , MAKENO_LM ");
                        strSql.AppendLine("                   , M6_CVNAM ");
                        strSql.AppendLine("                   , M6_CVNAM_CD ");
                        strSql.AppendLine("                   , M6_ITNAM ");
                        strSql.AppendLine("                   , M6_BREAK_REASON ");
                        strSql.AppendLine("                   , M6_AMT ");
                        strSql.AppendLine("                   , M6_DEVICE ");
                        strSql.AppendLine("                   , M6_DEVICE_CD ");
                        strSql.AppendLine("                   , WDATE ");
                        strSql.AppendLine("                   , ENT_DT ");
                        strSql.AppendLine("                   , ENT_ID ");
                        strSql.AppendLine("                   , MFY_DT ");
                        strSql.AppendLine("                   , MFY_ID ");
                        strSql.AppendLine("                   ) ");
                        strSql.AppendLine("              VALUES ");
                        strSql.AppendLine("                   ( " + sMakeNo + " ");
                        strSql.AppendLine("                   , " + dMakeNoLm + "");
                        strSql.AppendLine("                   , '" + sM6CvNam + "' ");
                        strSql.AppendLine("                   , '" + sM6CvNamCd + "' ");
                        strSql.AppendLine("                   , '" + sM6ItNam + "' ");
                        strSql.AppendLine("                   , '" + sM6BreakReason + "' ");
                        strSql.AppendLine("                   , " + dM6Amt + " ");
                        strSql.AppendLine("                   , '" + sM6Device + "' ");
                        strSql.AppendLine("                   , '" + sM6DeviceCd + "' ");
                        strSql.AppendLine("                   , '" + sWDate + "' ");
                        strSql.AppendLine("                   , CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("                   , '" + sEntId + "' ");
                        strSql.AppendLine("                   , CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("                   , '" + sMfyId + "') ");
                        strSql.AppendLine("   END");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }
                    Cursor = Cursors.Default;
                }

                #endregion[비용 저장부분]

                #region[검수 저장부분]

                DataTable dtGumsu = (DataTable)GridGumsu.DataSource;
                for(int i = 0; i < dtGumsu.Rows.Count; i++)
                {
                    string sEntId = dtGumsu.Rows[i]["ENT_ID"]?.ToString();
                    if (!FmMainToolBar2.UserID.Equals(sEntId))
                        continue;

                    string sMAKENO = dtGumsu.Rows[i]["MAKENO"]?.ToString();
                    string sMAKENO_LM = dtGumsu.Rows[i]["MAKENO_LM"]?.ToString();
                    string sM4_WGT_ADMT = dtGumsu.Rows[i]["M4_WGT_ADMT"]?.ToString();
                    string sM4_ITNL_YN = dtGumsu.Rows[i]["M4_ITNL_YN"]?.ToString();
                    string sM4_ISPT_OPN = dtGumsu.Rows[i]["M4_ISPT_OPN"]?.ToString();

                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" UPDATE MAKE_4 ");
                    strSql.AppendLine("    SET M4_WGT_ADMT = @M4_WGT_ADMT ");
                    strSql.AppendLine("      , M4_ITNL_YN = @M4_ITNL_YN ");
                    strSql.AppendLine("      , M4_ISPT_OPN = @M4_ISPT_OPN ");
                    strSql.AppendLine("  WHERE MAKENO = @MAKENO ");
                    strSql.AppendLine("    AND MAKENO_LM = @MAKENO_LM ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.Parameters.Clear();
                    cmd.Parameters.AddWithValue("@M4_WGT_ADMT", sM4_WGT_ADMT);
                    cmd.Parameters.AddWithValue("@M4_ITNL_YN", sM4_ITNL_YN);
                    cmd.Parameters.AddWithValue("@M4_ISPT_OPN", sM4_ISPT_OPN);
                    cmd.Parameters.AddWithValue("@MAKENO", sMAKENO);
                    cmd.Parameters.AddWithValue("@MAKENO_LM", sMAKENO_LM);
                    cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }

                #region 주석
                //if (GridViewGumsu.RowCount > 0)
                //{
                //    Cursor = Cursors.WaitCursor;

                //    DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridGumsu.DataSource);
                //    DataTable dtMerge = dsSave.Tables[0];

                //    for (int i = 0; i < dtMerge.Rows.Count; i++)
                //    {
                //        if (i == 0 || i == dtMerge.Rows.Count - 1)
                //        {
                //            string chksM4CarNo = dtMerge.Rows[i]["M4_CARNO"]?.ToString();
                //            string chksM4CvNam = dtMerge.Rows[i]["M4_CVNAM"]?.ToString();
                //            string chksM4CvNamIdtNo = dtMerge.Rows[i]["M4_CVNAM_IDTNO"].ToString();
                //            string chksM4Grade = dtMerge.Rows[i]["M4_GRADE"].ToString();
                //            string chksM4GradeCd = dtMerge.Rows[i]["M4_GRADE_CD"].ToString();
                //            string chksM4Wgt = dtMerge.Rows[i]["M4_WGT"].ToString();
                //            string chksM4Minus = dtMerge.Rows[i]["M4_MINUS"].ToString();
                //            string chksM4Evidence = dtMerge.Rows[i]["M4_EVIDENCE"].ToString();

                //            if (string.IsNullOrEmpty(chksM4CarNo) & string.IsNullOrEmpty(chksM4CvNam)
                //              & string.IsNullOrEmpty(chksM4CvNamIdtNo) & string.IsNullOrEmpty(chksM4Grade)
                //              & string.IsNullOrEmpty(chksM4GradeCd) & string.IsNullOrEmpty(chksM4Wgt)
                //              & string.IsNullOrEmpty(chksM4Minus) & string.IsNullOrEmpty(chksM4Evidence))
                //            {
                //                continue;
                //            }
                //        }

                //        string sMakeNo = dMakeNo.ToString();
                //        double dMakeNoLm = 0;

                //        if (String.IsNullOrEmpty(dtMerge.Rows[i]["MAKENO_LM"].ToString()))
                //        {
                //            strSql.Clear();
                //            strSql.AppendLine(" SELECT MAX(A.MAKENO_LM)");
                //            strSql.AppendLine("   FROM MAKE_4 A ");
                //            strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");

                //            cmd.CommandType = CommandType.Text;
                //            cmd.CommandText = strSql.ToString();

                //            if (String.IsNullOrEmpty(cmd.ExecuteScalar().ToString()))
                //            {
                //                dMakeNoLm = 1;
                //            }
                //            else
                //            {
                //                dMakeNoLm = Convert.ToDouble(cmd.ExecuteScalar()) + 1;
                //            }
                //        }
                //        else
                //        {
                //            dMakeNoLm = Convert.ToDouble(dtMerge.Rows[i]["MAKENO_LM"]);
                //        }

                //        string sM4CarNo = dtMerge.Rows[i]["M4_CARNO"].ToString();
                //        string sM4CvNam = dtMerge.Rows[i]["M4_CVNAM"].ToString();
                //        string sM4CvNamIdtNo = dtMerge.Rows[i]["M4_CVNAM_IDTNO"].ToString();
                //        string sM4Grade = dtMerge.Rows[i]["M4_GRADE"].ToString();
                //        string sM4GradeCd = dtMerge.Rows[i]["M4_GRADE_CD"].ToString();
                //        double dM4Wgt = String.IsNullOrEmpty(dtMerge.Rows[i]["M4_WGT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M4_WGT"]);
                //        double dM4Minus = String.IsNullOrEmpty(dtMerge.Rows[i]["M4_MINUS"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["M4_MINUS"]);
                //        string sM4Evidence = dtMerge.Rows[i]["M4_EVIDENCE"].ToString();
                //        string sEntId = FmMainToolBar2.UserID;
                //        string sMfyId = FmMainToolBar2.UserID;

                //        strSql.Clear();
                //        strSql.AppendLine(" ");
                //        strSql.AppendLine(" INSERT INTO MAKE_4 ");
                //        strSql.AppendLine("           ( MAKENO ");
                //        strSql.AppendLine("           , MAKENO_LM ");
                //        strSql.AppendLine("           , M4_CARNO ");
                //        strSql.AppendLine("           , M4_CVNAM ");
                //        strSql.AppendLine("           , M4_CVNAM_IDTNO ");
                //        strSql.AppendLine("           , M4_GRADE ");
                //        strSql.AppendLine("           , M4_GRADE_CD ");
                //        strSql.AppendLine("           , M4_WGT ");
                //        strSql.AppendLine("           , M4_MINUS ");
                //        strSql.AppendLine("           , M4_EVIDENCE ");
                //        strSql.AppendLine("           , ENT_DT ");
                //        strSql.AppendLine("           , ENT_ID ");
                //        strSql.AppendLine("           , MFY_DT ");
                //        strSql.AppendLine("           , MFY_ID ");
                //        strSql.AppendLine("           ) ");
                //        strSql.AppendLine("      VALUES ");
                //        strSql.AppendLine("           ( " + sMakeNo + " ");
                //        strSql.AppendLine("           , " + dMakeNoLm + "");
                //        strSql.AppendLine("           , '" + sM4CarNo + "' ");
                //        strSql.AppendLine("           , '" + sM4CvNam + "' ");
                //        strSql.AppendLine("           , '" + sM4CvNamIdtNo + "' ");
                //        strSql.AppendLine("           , '" + sM4Grade + "' ");
                //        strSql.AppendLine("           , '" + sM4GradeCd + "' ");
                //        strSql.AppendLine("           , " + dM4Wgt + " ");
                //        strSql.AppendLine("           , " + dM4Minus + " ");
                //        strSql.AppendLine("           , '" + sM4Evidence + "' ");
                //        strSql.AppendLine("           , NOW() ");
                //        strSql.AppendLine("           , '" + sEntId + "' ");
                //        strSql.AppendLine("           , NOW() ");
                //        strSql.AppendLine("           , '" + sMfyId + "') ");
                //        strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                //        strSql.AppendLine("             M4_CARNO = '" + sM4CarNo + "' ");
                //        strSql.AppendLine("           , M4_CVNAM = '" + sM4CvNam + "' ");
                //        strSql.AppendLine("           , M4_CVNAM_IDTNO = '" + sM4CvNamIdtNo + "' ");
                //        strSql.AppendLine("           , M4_GRADE = '" + sM4Grade + "' ");
                //        strSql.AppendLine("           , M4_GRADE_CD = '" + sM4GradeCd + "' ");
                //        strSql.AppendLine("           , M4_WGT = " + dM4Wgt + " ");
                //        strSql.AppendLine("           , M4_MINUS = " + dM4Minus + " ");
                //        strSql.AppendLine("           , M4_EVIDENCE = '" + sM4Evidence + "' ");
                //        strSql.AppendLine("           , MFY_DT = NOW() ");
                //        strSql.AppendLine("           , MFY_ID = '" + sMfyId + "' ");

                //        cmd.CommandType = CommandType.Text;
                //        cmd.CommandText = strSql.ToString();
                //        cmd.ExecuteNonQuery();
                //    }
                //    Cursor = Cursors.Default;
                //}
                #endregion

                #endregion[검수 저장부분]

                #region[설비점검 저장부분]
                string sMakeNo7 = dMakeNo.ToString();

                string sWorkTime = TimeEditWork.EditValue?.ToString();
                DateTime dtWorkTime = DateTime.Parse(sWorkTime);
                sWorkTime = dtWorkTime.ToString("HHmmss");

                //if (!string.IsNullOrEmpty(sWorkTime) && sWorkTime.Length == 19)
                //{
                //    sWorkTime = sWorkTime.Replace("-", "").Replace(":", "").Replace(" ", "").Substring(8, 6);
                //}

                string sM7_Chk01_A = string.Empty;
                string sM7_Chk01_B = string.Empty;
                string sM7_Chk02_A = string.Empty;
                string sM7_Chk02_B = string.Empty;
                string sM7_Chk03_A = string.Empty;
                string sM7_Chk03_B = string.Empty;
                string sM7_Chk04_A = string.Empty;
                string sM7_Chk04_B = string.Empty;
                string sM7_Chk05_A = string.Empty;
                string sM7_Chk05_B = string.Empty;
                string sM7_Chk06_A = string.Empty;
                string sM7_Chk06_B = string.Empty;
                string sM7_Chk07_A = string.Empty;
                string sM7_Chk07_B = string.Empty;
                string sM7_Chk08_A = string.Empty;
                string sM7_Chk08_B = string.Empty;
                string sM7_Chk09_A = string.Empty;
                string sM7_Chk09_B = string.Empty;
                string sM7_Chk10_A = string.Empty;
                string sM7_Chk10_B = string.Empty;
                string sM7_Chk11_A = string.Empty;
                string sM7_Chk11_B = string.Empty;
                string sM7_Chk12_A = string.Empty;
                string sM7_Chk12_B = string.Empty;
                string sM7_Chk13_A = string.Empty;
                string sM7_Chk13_B = string.Empty;
                string sM7_Chk14_A = string.Empty;
                string sM7_Chk14_B = string.Empty;

                string sM7_Chk15_A = string.Empty;
                string sM7_Chk15_B = string.Empty;
                string sM7_Chk16_A = string.Empty;
                string sM7_Chk16_B = string.Empty;

                string sEntId7 = FmMainToolBar2.UserID;
                string sMfyId7 = FmMainToolBar2.UserID;

                DataTable dtInspect = (DataTable)GridInspect.DataSource;
                dtInspect.TableName = "INSPECT";
                DataTable dtWeek = (DataTable)GridWeek.DataSource;
                dtWeek.TableName = "WEEK";
                DataTable dtMonth = (DataTable)GridMonth.DataSource;
                dtMonth.TableName = "MONTH";

                string[] arrBreakYn = new string[7];
                string[] arrBreakDesc = new string[7];
                for (int i = 0; i < 7 ; i++)
                {
                    if (i < 7)
                    {
                        arrBreakYn[i] = dtInspect.Rows[i][2].ToString();
                        arrBreakDesc[i] = dtInspect.Rows[i][3].ToString();
                    }
                }

                string[] arrBreakYnWeek = new string[7];
                string[] arrBreakDescWeek = new string[7];
                for (int i = 0; i < 7 ; i++)
                {
                    if (i < 7)
                    {
                        arrBreakYnWeek[i] = dtWeek.Rows[i][2].ToString();
                        arrBreakDescWeek[i] = dtWeek.Rows[i][3].ToString();
                    }
                }

                string[] arrBreakYnMonth = new string[7];
                string[] arrBreakDescMonth = new string[7];
                for (int i = 0; i < 2; i++)
                {
                    if (i < 2)
                    {
                        arrBreakYnMonth[i] = dtMonth.Rows[i][2].ToString();
                        arrBreakDescMonth[i] = dtMonth.Rows[i][3].ToString();
                    }
                }

                sM7_Chk01_A = arrBreakYn[0].ToString();
                sM7_Chk01_B = arrBreakDesc[0].ToString();
                sM7_Chk02_A = arrBreakYn[1].ToString();
                sM7_Chk02_B = arrBreakDesc[1].ToString();
                sM7_Chk03_A = arrBreakYn[2].ToString();
                sM7_Chk03_B = arrBreakDesc[2].ToString();
                sM7_Chk04_A = arrBreakYn[3].ToString();
                sM7_Chk04_B = arrBreakDesc[3].ToString();
                sM7_Chk05_A = arrBreakYn[4].ToString();
                sM7_Chk05_B = arrBreakDesc[4].ToString();
                sM7_Chk06_A = arrBreakYn[5].ToString();
                sM7_Chk06_B = arrBreakDesc[5].ToString();
                sM7_Chk07_A = arrBreakYn[6].ToString();
                sM7_Chk07_B = arrBreakDesc[6].ToString();
                sM7_Chk08_A = arrBreakYnWeek[0].ToString();
                sM7_Chk08_B = arrBreakDescWeek[0].ToString();
                sM7_Chk09_A = arrBreakYnWeek[1].ToString();
                sM7_Chk09_B = arrBreakDescWeek[1].ToString();
                sM7_Chk10_A = arrBreakYnWeek[2].ToString();
                sM7_Chk10_B = arrBreakDescWeek[2].ToString();
                sM7_Chk11_A = arrBreakYnWeek[3].ToString();
                sM7_Chk11_B = arrBreakDescWeek[3].ToString();
                sM7_Chk12_A = arrBreakYnWeek[4].ToString();
                sM7_Chk12_B = arrBreakDescWeek[4].ToString();
                sM7_Chk13_A = arrBreakYnWeek[5].ToString();
                sM7_Chk13_B = arrBreakDescWeek[5].ToString();
                sM7_Chk14_A = arrBreakYnWeek[6].ToString();
                sM7_Chk14_B = arrBreakDescWeek[6].ToString();
                sM7_Chk15_A = arrBreakYnMonth[0].ToString();
                sM7_Chk15_B = arrBreakDescMonth[0].ToString();
                sM7_Chk16_A = arrBreakYnMonth[1].ToString();
                sM7_Chk16_B = arrBreakDescMonth[1].ToString();

                strSql.Clear();
                #region MARIADB
                //strSql.AppendLine(" INSERT INTO MAKE_7 ");
                //strSql.AppendLine("           ( MAKENO ");
                //strSql.AppendLine("           , M7_MTNC_TIME ");
                //strSql.AppendLine("           , M7_CHECK01_A ");
                //strSql.AppendLine("           , M7_CHECK01_B ");
                //strSql.AppendLine("           , M7_CHECK02_A ");
                //strSql.AppendLine("           , M7_CHECK02_B ");
                //strSql.AppendLine("           , M7_CHECK03_A ");
                //strSql.AppendLine("           , M7_CHECK03_B ");
                //strSql.AppendLine("           , M7_CHECK04_A ");
                //strSql.AppendLine("           , M7_CHECK04_B ");
                //strSql.AppendLine("           , M7_CHECK05_A ");
                //strSql.AppendLine("           , M7_CHECK05_B ");
                //strSql.AppendLine("           , M7_CHECK06_A ");
                //strSql.AppendLine("           , M7_CHECK06_B ");
                //strSql.AppendLine("           , M7_CHECK07_A ");
                //strSql.AppendLine("           , M7_CHECK07_B ");
                //strSql.AppendLine("           , M7_CHECK08_A ");
                //strSql.AppendLine("           , M7_CHECK08_B ");
                //strSql.AppendLine("           , M7_CHECK09_A ");
                //strSql.AppendLine("           , M7_CHECK09_B ");
                //strSql.AppendLine("           , M7_CHECK10_A ");
                //strSql.AppendLine("           , M7_CHECK10_B ");
                //strSql.AppendLine("           , M7_CHECK11_A ");
                //strSql.AppendLine("           , M7_CHECK11_B ");
                //strSql.AppendLine("           , M7_CHECK12_A ");
                //strSql.AppendLine("           , M7_CHECK12_B ");
                //strSql.AppendLine("           , M7_CHECK13_A ");
                //strSql.AppendLine("           , M7_CHECK13_B ");
                //strSql.AppendLine("           , M7_CHECK14_A ");
                //strSql.AppendLine("           , M7_CHECK14_B ");
                //strSql.AppendLine("           , ENT_DT ");
                //strSql.AppendLine("           , ENT_ID ");
                //strSql.AppendLine("           , MFY_DT ");
                //strSql.AppendLine("           , MFY_ID )");
                //strSql.AppendLine("      VALUES ");
                //strSql.AppendLine("           ( " + sMakeNo7 + " ");
                //strSql.AppendLine("           , '" + sWorkTime + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk01_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk01_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk02_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk02_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk03_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk03_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk04_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk04_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk05_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk05_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk06_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk06_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk07_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk07_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk08_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk08_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk09_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk09_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk10_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk10_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk11_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk11_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk12_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk12_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk13_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk13_B + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk14_A + "' ");
                //strSql.AppendLine("           , '" + sM7_Chk14_B + "' ");
                //strSql.AppendLine("           , NOW() ");
                //strSql.AppendLine("           , '" + sEntId7 + "' ");
                //strSql.AppendLine("           , NOW() ");
                //strSql.AppendLine("           , '" + sMfyId7 + "' )");
                //strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                //strSql.AppendLine("             M7_MTNC_TIME = '" + sWorkTime + "' ");
                //strSql.AppendLine("           , M7_CHECK01_A = '" + sM7_Chk01_A + "' ");
                //strSql.AppendLine("           , M7_CHECK01_B = '" + sM7_Chk01_B + "' ");
                //strSql.AppendLine("           , M7_CHECK02_A = '" + sM7_Chk02_A + "' ");
                //strSql.AppendLine("           , M7_CHECK02_B = '" + sM7_Chk02_B + "' ");
                //strSql.AppendLine("           , M7_CHECK03_A = '" + sM7_Chk03_A + "' ");
                //strSql.AppendLine("           , M7_CHECK03_B = '" + sM7_Chk03_B + "' ");
                //strSql.AppendLine("           , M7_CHECK04_A = '" + sM7_Chk04_A + "' ");
                //strSql.AppendLine("           , M7_CHECK04_B = '" + sM7_Chk04_B + "' ");
                //strSql.AppendLine("           , M7_CHECK05_A = '" + sM7_Chk05_A + "' ");
                //strSql.AppendLine("           , M7_CHECK05_B = '" + sM7_Chk05_B + "' ");
                //strSql.AppendLine("           , M7_CHECK06_A = '" + sM7_Chk06_A + "' ");
                //strSql.AppendLine("           , M7_CHECK06_B = '" + sM7_Chk06_B + "' ");
                //strSql.AppendLine("           , M7_CHECK07_A = '" + sM7_Chk07_A + "' ");
                //strSql.AppendLine("           , M7_CHECK07_B = '" + sM7_Chk07_B + "' ");
                //strSql.AppendLine("           , M7_CHECK08_A = '" + sM7_Chk08_A + "' ");
                //strSql.AppendLine("           , M7_CHECK08_B = '" + sM7_Chk08_B + "' ");
                //strSql.AppendLine("           , M7_CHECK09_A = '" + sM7_Chk09_A + "' ");
                //strSql.AppendLine("           , M7_CHECK09_B = '" + sM7_Chk09_B + "' ");
                //strSql.AppendLine("           , M7_CHECK10_A = '" + sM7_Chk10_A + "' ");
                //strSql.AppendLine("           , M7_CHECK10_B = '" + sM7_Chk10_B + "' ");
                //strSql.AppendLine("           , M7_CHECK11_A = '" + sM7_Chk11_A + "' ");
                //strSql.AppendLine("           , M7_CHECK11_B = '" + sM7_Chk11_B + "' ");
                //strSql.AppendLine("           , M7_CHECK12_A = '" + sM7_Chk12_A + "' ");
                //strSql.AppendLine("           , M7_CHECK12_B = '" + sM7_Chk12_B + "' ");
                //strSql.AppendLine("           , M7_CHECK13_A = '" + sM7_Chk13_A + "' ");
                //strSql.AppendLine("           , M7_CHECK13_B = '" + sM7_Chk13_B + "' ");
                //strSql.AppendLine("           , M7_CHECK14_A = '" + sM7_Chk14_A + "' ");
                //strSql.AppendLine("           , M7_CHECK14_B = '" + sM7_Chk14_B + "' ");
                //strSql.AppendLine("           , MFY_DT = NOW() ");
                //strSql.AppendLine("           , MFY_ID = '" + sMfyId7 + "' ");
                #endregion
                strSql.AppendLine("IF EXISTS(SELECT MAKENO FROM MAKE_7 WHERE MAKENO = "+ sMakeNo7 + ") ");
                strSql.AppendLine("    BEGIN                                                     ");
                strSql.AppendLine("          UPDATE MAKE_7                                       ");
                strSql.AppendLine("             SET M7_MTNC_TIME = '" + sWorkTime + "' ");
                strSql.AppendLine("               , M7_CHECK01_A = '" + sM7_Chk01_A + "' ");
                strSql.AppendLine("               , M7_CHECK01_B = '" + sM7_Chk01_B + "' ");
                strSql.AppendLine("               , M7_CHECK02_A = '" + sM7_Chk02_A + "' ");
                strSql.AppendLine("               , M7_CHECK02_B = '" + sM7_Chk02_B + "' ");
                strSql.AppendLine("               , M7_CHECK03_A = '" + sM7_Chk03_A + "' ");
                strSql.AppendLine("               , M7_CHECK03_B = '" + sM7_Chk03_B + "' ");
                strSql.AppendLine("               , M7_CHECK04_A = '" + sM7_Chk04_A + "' ");
                strSql.AppendLine("               , M7_CHECK04_B = '" + sM7_Chk04_B + "' ");
                strSql.AppendLine("               , M7_CHECK05_A = '" + sM7_Chk05_A + "' ");
                strSql.AppendLine("               , M7_CHECK05_B = '" + sM7_Chk05_B + "' ");
                strSql.AppendLine("               , M7_CHECK06_A = '" + sM7_Chk06_A + "' ");
                strSql.AppendLine("               , M7_CHECK06_B = '" + sM7_Chk06_B + "' ");
                strSql.AppendLine("               , M7_CHECK07_A = '" + sM7_Chk07_A + "' ");
                strSql.AppendLine("               , M7_CHECK07_B = '" + sM7_Chk07_B + "' ");
                strSql.AppendLine("               , M7_CHECK08_A = '" + sM7_Chk08_A + "' ");
                strSql.AppendLine("               , M7_CHECK08_B = '" + sM7_Chk08_B + "' ");
                strSql.AppendLine("               , M7_CHECK09_A = '" + sM7_Chk09_A + "' ");
                strSql.AppendLine("               , M7_CHECK09_B = '" + sM7_Chk09_B + "' ");
                strSql.AppendLine("               , M7_CHECK10_A = '" + sM7_Chk10_A + "' ");
                strSql.AppendLine("               , M7_CHECK10_B = '" + sM7_Chk10_B + "' ");
                strSql.AppendLine("               , M7_CHECK11_A = '" + sM7_Chk11_A + "' ");
                strSql.AppendLine("               , M7_CHECK11_B = '" + sM7_Chk11_B + "' ");
                strSql.AppendLine("               , M7_CHECK12_A = '" + sM7_Chk12_A + "' ");
                strSql.AppendLine("               , M7_CHECK12_B = '" + sM7_Chk12_B + "' ");
                strSql.AppendLine("               , M7_CHECK13_A = '" + sM7_Chk13_A + "' ");
                strSql.AppendLine("               , M7_CHECK13_B = '" + sM7_Chk13_B + "' ");
                strSql.AppendLine("               , M7_CHECK14_A = '" + sM7_Chk14_A + "' ");
                strSql.AppendLine("               , M7_CHECK14_B = '" + sM7_Chk14_B + "' ");
                strSql.AppendLine("               , M7_CHECK15_A = '" + sM7_Chk15_A + "' ");
                strSql.AppendLine("               , M7_CHECK15_B = '" + sM7_Chk15_B + "' ");
                strSql.AppendLine("               , M7_CHECK16_A = '" + sM7_Chk16_A + "' ");
                strSql.AppendLine("               , M7_CHECK16_B = '" + sM7_Chk16_B + "' ");
                strSql.AppendLine("               , MFY_DT = CONVERT(VARCHAR(20), GETDATE(), 20)    ");
                strSql.AppendLine("               , MFY_ID = '" + sMfyId7 + "' ");
                strSql.AppendLine("           WHERE MAKENO = " + sMakeNo7 + "");
                strSql.AppendLine("    END                                                       ");
                strSql.AppendLine("    ELSE                                                      ");
                strSql.AppendLine("    BEGIN                                                     ");
                strSql.AppendLine("          INSERT INTO MAKE_7                                  ");
                strSql.AppendLine("               (MAKENO                                        ");
                strSql.AppendLine("               , M7_MTNC_TIME                                 ");
                strSql.AppendLine("               , M7_CHECK01_A                                 ");
                strSql.AppendLine("               , M7_CHECK01_B                                 ");
                strSql.AppendLine("               , M7_CHECK02_A                                 ");
                strSql.AppendLine("               , M7_CHECK02_B                                 ");
                strSql.AppendLine("               , M7_CHECK03_A                                 ");
                strSql.AppendLine("               , M7_CHECK03_B                                 ");
                strSql.AppendLine("               , M7_CHECK04_A                                 ");
                strSql.AppendLine("               , M7_CHECK04_B                                 ");
                strSql.AppendLine("               , M7_CHECK05_A                                 ");
                strSql.AppendLine("               , M7_CHECK05_B                                 ");
                strSql.AppendLine("               , M7_CHECK06_A                                 ");
                strSql.AppendLine("               , M7_CHECK06_B                                 ");
                strSql.AppendLine("               , M7_CHECK07_A                                 ");
                strSql.AppendLine("               , M7_CHECK07_B                                 ");
                strSql.AppendLine("               , M7_CHECK08_A                                 ");
                strSql.AppendLine("               , M7_CHECK08_B                                 ");
                strSql.AppendLine("               , M7_CHECK09_A                                 ");
                strSql.AppendLine("               , M7_CHECK09_B                                 ");
                strSql.AppendLine("               , M7_CHECK10_A                                 ");
                strSql.AppendLine("               , M7_CHECK10_B                                 ");
                strSql.AppendLine("               , M7_CHECK11_A                                 ");
                strSql.AppendLine("               , M7_CHECK11_B                                 ");
                strSql.AppendLine("               , M7_CHECK12_A                                 ");
                strSql.AppendLine("               , M7_CHECK12_B                                 ");
                strSql.AppendLine("               , M7_CHECK13_A                                 ");
                strSql.AppendLine("               , M7_CHECK13_B                                 ");
                strSql.AppendLine("               , M7_CHECK14_A                                 ");
                strSql.AppendLine("               , M7_CHECK14_B                                 ");
                strSql.AppendLine("               , M7_CHECK15_A                                 ");
                strSql.AppendLine("               , M7_CHECK15_B                                 ");
                strSql.AppendLine("               , M7_CHECK16_A                                 ");
                strSql.AppendLine("               , M7_CHECK16_B                                 ");
                strSql.AppendLine("               , ENT_DT                                       ");
                strSql.AppendLine("               , ENT_ID                                       ");
                strSql.AppendLine("               , MFY_DT                                       ");
                strSql.AppendLine("               , MFY_ID )                                     ");
                strSql.AppendLine("          VALUES                                              ");
                strSql.AppendLine("           ( " + sMakeNo7 + " ");
                strSql.AppendLine("           , '" + sWorkTime + "' ");
                strSql.AppendLine("           , '" + sM7_Chk01_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk01_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk02_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk02_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk03_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk03_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk04_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk04_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk05_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk05_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk06_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk06_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk07_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk07_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk08_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk08_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk09_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk09_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk10_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk10_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk11_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk11_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk12_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk12_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk13_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk13_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk14_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk14_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk15_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk15_B + "' ");
                strSql.AppendLine("           , '" + sM7_Chk16_A + "' ");
                strSql.AppendLine("           , '" + sM7_Chk16_B + "' ");
                strSql.AppendLine("               , CONVERT(VARCHAR(20), GETDATE(), 20)             ");
                strSql.AppendLine("           , '" + sEntId7 + "' ");
                strSql.AppendLine("               , CONVERT(VARCHAR(20), GETDATE(), 20)             ");
                strSql.AppendLine("           , '" + sMfyId7 + "' )");
                strSql.AppendLine("    END                                                       ");

                #endregion[설비점검 저장부분]

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                Dispose();

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }
        
        private void BtnExport_Click(object sender, EventArgs e)
        {

        }

        private void ProdMgtReport_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {

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
                
            }
            else if (e.KeyCode == Keys.F8)
            {

            }
        }
        
        private void BtnClose_Click(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                
                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT COUNT(CASE WHEN A.CNT > 0 THEN 1 END ) AS EXIST_CNT ");
                strSql.AppendLine("   FROM ( ");
                strSql.AppendLine("        SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("          FROM MAKE_1 X1  ");
                strSql.AppendLine("         WHERE X1.MAKENO = " + dMakeNo + " ");
                strSql.AppendLine("         UNION ALL  ");
                strSql.AppendLine("        SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("          FROM MAKE_2 X1  ");
                strSql.AppendLine("         WHERE X1.MAKENO = " + dMakeNo + " ");
                strSql.AppendLine("         UNION ALL  ");
                strSql.AppendLine("        SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("          FROM MAKE_3 X1  ");
                strSql.AppendLine("         WHERE X1.MAKENO = " + dMakeNo + " ");
                strSql.AppendLine("         UNION ALL  ");
                strSql.AppendLine("        SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("          FROM MAKE_4 X1  ");
                strSql.AppendLine("         WHERE X1.MAKENO = " + dMakeNo + " ");
                strSql.AppendLine("         UNION ALL  ");
                strSql.AppendLine("        SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("          FROM MAKE_5 X1  ");
                strSql.AppendLine("         WHERE X1.MAKENO = " + dMakeNo + " ");
                strSql.AppendLine("         UNION ALL  ");
                strSql.AppendLine("        SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("          FROM MAKE_6 X1  ");
                strSql.AppendLine("         WHERE X1.MAKENO = " + dMakeNo + " ");
                strSql.AppendLine("        UNION ALL  ");
                strSql.AppendLine("        SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("          FROM MAKE_7 X1  ");
                strSql.AppendLine("         WHERE X1.MAKENO = " + dMakeNo + " ");
                strSql.AppendLine("        UNION ALL  ");
                strSql.AppendLine("        SELECT COUNT(*) AS CNT ");
                strSql.AppendLine("          FROM MAKE_EXPENSE X1  ");
                strSql.AppendLine("         WHERE X1.MAKENO = " + dMakeNo + " ");
                strSql.AppendLine("  ) A ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                if (dt != null && dt.Rows.Count > 0)
                {
                    int iCnt = Convert.ToInt32(dt.Rows[0]["EXIST_CNT"]);
                    if (iCnt == 0)
                    {
                        DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                        SqlCommand cmd = DBConn.dbCon.CreateCommand();
                        cmd.Transaction = DBConn.dbTran;

                        strSql.Clear();
                        strSql.AppendLine(" DELETE FROM MAKE_M ");
                        strSql.AppendLine(" WHERE MAKENO = " + dMakeNo + " ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        DBConn.dbTran.Commit();
                        DBConn.dbTran = null;
                    }
                }
                Dispose();
                Cursor = Cursors.Default;
            }
            catch(Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
                Cursor = Cursors.Default;
            }

            Cursor = Cursors.Default;
            
        }

        private void BtnCostAdder_Click(object sender, EventArgs e)
        {
            if (CheckApprovalYN())
            {
                XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT MAX(A.MAKENO_LN) AS MAX_VALUE ");
            strSql.AppendLine("   FROM MAKE_EXPENSE A ");
            strSql.AppendLine("  WHERE MAKENO = " + dMakeNo + " ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            double dMakeNoLn = 0;
            if(dt.Rows.Count == 0)
            {
                dMakeNoLn = 0;
            }
            else
            {
                string sTemp = dt.Rows[0]["MAX_VALUE"]?.ToString();
                if (string.IsNullOrEmpty(sTemp))
                {
                    dMakeNoLn = 0;
                }
                else
                {
                    dMakeNoLn = Convert.ToDouble(sTemp) + 1;
                }
            }

            string sDate = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            ProdCostAdder frm = new ProdCostAdder();
            frm.sMakeNo = dMakeNo.ToString();
            frm.sMakeNoLn = dMakeNoLn.ToString();
            frm.sEmpID = drUserInfo["EMP_ID"]?.ToString();
            frm.ProcessDate = sDate;
            frm.sEmpNm = drUserInfo["EMP_NM"].ToString();

            if (frm.ShowDialog() == DialogResult.OK)
            {
                GetMake6Info(dMakeNo.ToString());
            }
        }

        private void GridViewCost_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                string MAkeNoLn = GridViewCost.GetFocusedRowCellValue("MAKENO_LN").ToString();

                ProdCostAdder frm = new ProdCostAdder();
                frm.sMakeNo = sMakeNo;
                frm.sMakeNoLn = MAkeNoLn;
                frm.sEmpID = drUserInfo["EMP_ID"]?.ToString();

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    GetMake6Info(sMakeNo);
                }
            }
        }

        private bool CheckApprovalYN()
        {
            string WorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sEmpCd = drUserInfo["EMP_ID"]?.ToString();
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.REAL_DUTY_DEPT ");
            strSql.AppendLine("   FROM HR_EMP_BASIS A ");
            strSql.AppendLine("  WHERE EMP_ID = '" + sEmpCd + "' ");
            strSql.AppendLine("    AND EMPL_GB = 'Y' ");
            DataTable dtDept = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            string sDept = dtDept.Rows[0]["REAL_DUTY_DEPT"]?.ToString();


            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT  ");
                if (sDept == "4100") { strSql.AppendLine("       A.SIGN1  AS Sign"); }
                if (sDept == "4150") { strSql.AppendLine("       A.SIGN1a  AS Sign"); }
                if (sDept == "4200") { strSql.AppendLine("       A.SIGN2  AS Sign"); }
                if (sDept == "4300") { strSql.AppendLine("       A.SIGN3  AS Sign"); }
            strSql.AppendLine("   FROM MAKE_S A ");
            strSql.AppendLine("  WHERE MDATE = '" + WorkYmd + "' ");
            strSql.AppendLine("    AND GUBUN = '1' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            
            if(dt != null)
            {
                if(dt.Rows.Count > 0)
                {
                    if (dt.Rows[0]["Sign"].ToString().Equals("Y"))
                    {
                        return true;
                    }
                    else if (dt.Rows[0]["Sign"].ToString().Equals("N"))
                    {
                        return false;
                    }
                }

                return false;
            }
            else
            {
                return false;
            }

        }

        private void DateEditYmd_Leave(object sender, EventArgs e)
        {
            //if (!sReadOnly)
            //{
            //    string sEmpId = drUserInfo["EMP_ID"]?.ToString();
            //    string sWorkYmd1 = DateEditYmd.EditValue?.ToString().Substring(0, 10);
            //    string sWorkYmd = DateEditYmd.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            //    StringBuilder strSql = new StringBuilder();

            //    strSql.AppendLine(" SELECT CASE WHEN COUNT(1) > 0 THEN 'Y' ELSE 'N' END AS EXIST_YN ");
            //    strSql.AppendLine("   FROM MAKE_M  ");
            //    strSql.AppendLine("  WHERE MDATE = '" + sWorkYmd + "'");
            //    strSql.AppendLine("    AND MUSER_ID = '" + sEmpId + "' ");

            //    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            //    string sYn = dt.Rows[0]["EXIST_YN"]?.ToString();
            //    if (sYn.Equals("Y"))
            //    {
            //        MessageBox.Show(sWorkYmd1 + "의 생산일보 내역이 존재합니다.\r\n일자 변경 후 생산일보를 작성해주십시오.");
            //    }
            //}
        }

        private void BtnClose_MouseHover(object sender, EventArgs e)
        {
            DateEditYmd.Leave -= DateEditYmd_Leave;
        }

        private void BtnClose_MouseLeave(object sender, EventArgs e)
        {
            //if(!sReadOnly) DateEditYmd.Leave += DateEditYmd_Leave;
        }

        private void ProdMgtReport_TextChanged(object sender, EventArgs e)
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

        private void GridViewWork_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void ProdMgtReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            BtnClose_Click(null, null);
        }

        private void RepoGumsuBtnImage_Click(object sender, EventArgs e)
        {
            
        }

        public GridView RST_GRID_VIEW;
        private void BtnGumsu_Click(object sender, EventArgs e)
        {
            PD01001F03 frm = new PD01001F03();
            frm.P_ProdMgtReport = this;
            frm.MAKENO = sMakeNo;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                SaveInspectInfo(RST_GRID_VIEW);
            }
        }

        private void SaveInspectInfo(GridView view)
        {
            try
            {
                Cursor = Cursors.WaitCursor;

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                Cursor = Cursors.Default;

                StringBuilder strSql = new StringBuilder();
                
                #region[검수 저장부분]
                if (view.SelectedRowsCount > 0)
                {
                    Cursor = Cursors.WaitCursor;

                    int[] iArrRowIdx = view.GetSelectedRows();

                    for (int i = 0; i < iArrRowIdx.Length; i++)
                    {
                        string sMakeNo = dMakeNo.ToString();
                        double dMakeNoLm = 0;

                        strSql.Clear();
                        strSql.AppendLine(" SELECT MAX(A.MAKENO_LM)");
                        strSql.AppendLine("   FROM MAKE_4 A ");
                        strSql.AppendLine("  WHERE MAKENO = '" + sMakeNo + "' ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();

                        if (string.IsNullOrEmpty(cmd.ExecuteScalar().ToString()))
                        {
                            dMakeNoLm = 1;
                        }
                        else
                        {
                            dMakeNoLm = Convert.ToDouble(cmd.ExecuteScalar()) + 1;
                        }

                        DataRow row = view.GetDataRow(iArrRowIdx[i]);

                        string sM4CarNo = row["CARNO"]?.ToString();
                        string sM4CvCod = row["DEALER_CD"]?.ToString();
                        string sM4CvNam = row["DEALER_NM"]?.ToString();
                        string sM4CvNamIdtNo = row["IDT_NO"]?.ToString();
                        string sM4Grade = row["GRADE_NM"]?.ToString();
                        string sM4GradeCd = row["GRADE_CD"]?.ToString();
                        double dM4Wgt = string.IsNullOrEmpty(row["DJ_WEIGHT"]?.ToString()) ? 0 : Convert.ToDouble(row["DJ_WEIGHT"]);
                        double dM4Minus = string.IsNullOrEmpty(row["LOSS"]?.ToString()) ? 0 : Convert.ToDouble(row["LOSS"]);
                        string sM4Evidence = row["ISPT_NOTE"]?.ToString();
                        string sEntId = FmMainToolBar2.UserID;
                        string sMfyId = FmMainToolBar2.UserID;

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        #region mariaDB
                        //strSql.AppendLine(" INSERT INTO MAKE_4 ");
                        //strSql.AppendLine("           ( MAKENO ");
                        //strSql.AppendLine("           , MAKENO_LM ");
                        //strSql.AppendLine("           , M4_CARNO ");
                        //strSql.AppendLine("           , M4_CVCOD ");
                        //strSql.AppendLine("           , M4_CVNAM ");
                        //strSql.AppendLine("           , M4_CVNAM_IDTNO ");
                        //strSql.AppendLine("           , M4_GRADE ");
                        //strSql.AppendLine("           , M4_GRADE_CD ");
                        //strSql.AppendLine("           , M4_WGT ");
                        //strSql.AppendLine("           , M4_MINUS ");
                        //strSql.AppendLine("           , M4_EVIDENCE ");
                        //strSql.AppendLine("           , ENT_DT ");
                        //strSql.AppendLine("           , ENT_ID ");
                        //strSql.AppendLine("           , MFY_DT ");
                        //strSql.AppendLine("           , MFY_ID ");
                        //strSql.AppendLine("           ) ");
                        //strSql.AppendLine("      VALUES ");
                        //strSql.AppendLine("           ( @MAKENO ");
                        //strSql.AppendLine("           , @MAKENO_LN ");
                        //strSql.AppendLine("           , '" + sM4CarNo + "' ");
                        //strSql.AppendLine("           , @M4_CVCOD ");
                        //strSql.AppendLine("           , '" + sM4CvNam + "' ");
                        //strSql.AppendLine("           , '" + sM4CvNamIdtNo + "' ");
                        //strSql.AppendLine("           , '" + sM4Grade + "' ");
                        //strSql.AppendLine("           , '" + sM4GradeCd + "' ");
                        //strSql.AppendLine("           , " + dM4Wgt + " ");
                        //strSql.AppendLine("           , " + dM4Minus + " ");
                        //strSql.AppendLine("           , '" + sM4Evidence + "' ");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sEntId + "' ");
                        //strSql.AppendLine("           , NOW() ");
                        //strSql.AppendLine("           , '" + sMfyId + "') ");
                        #endregion

                        strSql.AppendLine(" INSERT INTO MAKE_4 ");
                        strSql.AppendLine("           ( MAKENO ");
                        strSql.AppendLine("           , MAKENO_LM ");
                        strSql.AppendLine("           , M4_CARNO ");
                        strSql.AppendLine("           , M4_CVCOD ");
                        strSql.AppendLine("           , M4_CVNAM ");
                        strSql.AppendLine("           , M4_CVNAM_IDTNO ");
                        strSql.AppendLine("           , M4_GRADE ");
                        strSql.AppendLine("           , M4_GRADE_CD ");
                        strSql.AppendLine("           , M4_WGT ");
                        strSql.AppendLine("           , M4_MINUS ");
                        strSql.AppendLine("           , M4_EVIDENCE ");
                        strSql.AppendLine("           , ENT_DT ");
                        strSql.AppendLine("           , ENT_ID ");
                        strSql.AppendLine("           , MFY_DT ");
                        strSql.AppendLine("           , MFY_ID ");
                        strSql.AppendLine("           ) ");
                        strSql.AppendLine("      VALUES ");
                        strSql.AppendLine("           ( @MAKENO ");
                        strSql.AppendLine("           , @MAKENO_LN ");
                        strSql.AppendLine("           , '" + sM4CarNo + "' ");
                        strSql.AppendLine("           , @M4_CVCOD ");
                        strSql.AppendLine("           , '" + sM4CvNam + "' ");
                        strSql.AppendLine("           , '" + sM4CvNamIdtNo + "' ");
                        strSql.AppendLine("           , '" + sM4Grade + "' ");
                        strSql.AppendLine("           , '" + sM4GradeCd + "' ");
                        strSql.AppendLine("           , " + dM4Wgt + " ");
                        strSql.AppendLine("           , " + dM4Minus + " ");
                        strSql.AppendLine("           , '" + sM4Evidence + "' ");
                        strSql.AppendLine("           , CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("           , '" + sEntId + "' ");
                        strSql.AppendLine("           , CONVERT([varchar](20), getdate(), (21)) ");
                        strSql.AppendLine("           , '" + sMfyId + "') ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.Parameters.Clear();
                        cmd.Parameters.AddWithValue("@MAKENO", sMakeNo);
                        cmd.Parameters.AddWithValue("@MAKENO_LN", dMakeNoLm);
                        cmd.Parameters.AddWithValue("@M4_CVCOD", sM4CvCod);
                        cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                    }
                    Cursor = Cursors.Default;
                }
                #endregion[검수 저장부분]
                
                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                GetMake4Info(sMakeNo);

                Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void RepoGumsuBtnImage_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            string sMakeNo = GridViewGumsu.GetFocusedRowCellValue("MAKENO")?.ToString();
            string sMakeNo_Ln = GridViewGumsu.GetFocusedRowCellValue("MAKENO_LM")?.ToString();

            if (string.IsNullOrEmpty(sMakeNo) || string.IsNullOrEmpty(sMakeNo_Ln))
            {
                XtraMessageBox.Show("올바른 데이터를 선택하세요.");
                return;
            }

            PD01001F02 frm = new PD01001F02();

            frm.MAKENO = sMakeNo;
            frm.MAKENO_LN = sMakeNo_Ln;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                GetMake4Info(sMakeNo);
            }
        }
        
        private void XTabControl_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if (XTabControl.SelectedTabPageIndex == 1)
            {
                GetMake4Info(sMakeNo);
            }
        }

        private void BtnGumsuDel_Click(object sender, EventArgs e)
        {
            PD01001F04 frm = new PD01001F04();
            frm.MAKENO = sMakeNo;
            if(frm.ShowDialog() == DialogResult.OK)
            {
                GetMake4Info(sMakeNo);
            }
        }

        private void RepoTxtGumsuIsptOpn_EditValueChanging(object sender, DevExpress.XtraEditors.Controls.ChangingEventArgs e)
        {
            string sEntId = GridViewGumsu.GetFocusedRowCellValue(GridColGumsuEntId)?.ToString();
            string sEntNm = GridViewGumsu.GetFocusedRowCellValue(GridColGumsuEntNm)?.ToString();
            if (!FmMainToolBar2.UserID.Equals(sEntId))
            {
                string sMSG = string.Format("{0} 번째 행은 {1}님이 등록한 검수 건이므로 수정하실 수 없습니다. ", GridViewGumsu.FocusedRowHandle + 1, sEntNm);
                XtraMessageBox.Show(sMSG);
                GridViewGumsu.SetFocusedRowCellValue(GridColGumsuIsptOpn, e.OldValue);
                return;
            }
        }

        private void RepoTxtGumsuWgtAdmt_EditValueChanging(object sender, DevExpress.XtraEditors.Controls.ChangingEventArgs e)
        {
            string sEntId = GridViewGumsu.GetFocusedRowCellValue(GridColGumsuEntId)?.ToString();
            string sEntNm = GridViewGumsu.GetFocusedRowCellValue(GridColGumsuEntNm)?.ToString();
            if (!FmMainToolBar2.UserID.Equals(sEntId))
            {
                string sMSG = string.Format("{0} 번째 행은 {1}님이 등록한 검수 건이므로 수정하실 수 없습니다. ", GridViewGumsu.FocusedRowHandle + 1, sEntNm);
                XtraMessageBox.Show(sMSG);
                GridViewGumsu.SetFocusedRowCellValue(GridColGumsuWtAdmt, e.OldValue);
                return;
            }
        }

        private void RepoChkGumsuItnlYN_EditValueChanging(object sender, DevExpress.XtraEditors.Controls.ChangingEventArgs e)
        {
            string sEntId = GridViewGumsu.GetFocusedRowCellValue(GridColGumsuEntId)?.ToString();
            string sEntNm = GridViewGumsu.GetFocusedRowCellValue(GridColGumsuEntNm)?.ToString();
            if (!FmMainToolBar2.UserID.Equals(sEntId))
            {
                string sMSG = string.Format("{0} 번째 행은 {1}님이 등록한 검수 건이므로 수정하실 수 없습니다. ", GridViewGumsu.FocusedRowHandle + 1, sEntNm);
                XtraMessageBox.Show(sMSG);
                GridViewGumsu.SetFocusedRowCellValue(GridColGumsuItnlYn, e.OldValue);
                return;
            }
        }

        private void GridViewGumsu_ShowingEditor(object sender, CancelEventArgs e)
        {
            string sEntId = GridViewGumsu.GetFocusedRowCellValue(GridColGumsuEntId)?.ToString();
            if (!FmMainToolBar2.UserID.Equals(sEntId))
            {
                e.Cancel = true;
            }
        }

        private void ProdMgtReport_KeyDown_1(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
            {
                if (CheckApprovalYN())
                {
                    XtraMessageBox.Show("해당 건은 결제승인 상태입니다.");
                    return;
                }

                if (XtraMessageBox.Show("작성한 내용을 저장하시겠습니까?"
                    , "저장여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }
                else
                {
                    BtnSave.PerformClick();
                }

            }
        }

        private void RepoWorkTimeEdit_EditValueChanged(object sender, EventArgs e)
        {
            TimeEdit timeEdit = (TimeEdit)sender;

            string sDate = timeEdit.EditValue?.ToString();

            if(DateTime.TryParse(sDate, out DateTime dateTime))
            {
                GridViewWork.SetRowCellValue(GridViewWork.FocusedRowHandle, GridViewWork.FocusedColumn, dateTime.ToString("HH:mm"));
            }
            else
            {
                GridViewWork.SetRowCellValue(GridViewWork.FocusedRowHandle, GridViewWork.FocusedColumn, "");
            }
        }

        private void RepoGilloTimeEdit_EditValueChanged(object sender, EventArgs e)
        {
            TimeEdit timeEdit = (TimeEdit)sender;

            string sDate = timeEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dateTime))
            {
                GridViewGuillotine.SetRowCellValue(GridViewGuillotine.FocusedRowHandle, GridViewGuillotine.FocusedColumn, dateTime.ToString("HH:mm"));
            }
            else
            {
                GridViewGuillotine.SetRowCellValue(GridViewGuillotine.FocusedRowHandle, GridViewGuillotine.FocusedColumn, "");
            }
        }

        private void RepoMntncTimeEdit_EditValueChanged(object sender, EventArgs e)
        {
            TimeEdit timeEdit = (TimeEdit)sender;

            string sDate = timeEdit.EditValue?.ToString();

            if (DateTime.TryParse(sDate, out DateTime dateTime))
            {
                GridViewShreder.SetRowCellValue(GridViewShreder.FocusedRowHandle, GridViewShreder.FocusedColumn, dateTime.ToString("HH:mm"));
            }
            else
            {
                GridViewShreder.SetRowCellValue(GridViewShreder.FocusedRowHandle, GridViewShreder.FocusedColumn, "");
            }
        }

        private void RepoTxtShrederTime_EditValueChanged(object sender, EventArgs e)//스크류가동
        {
            TextEdit textEdit = (TextEdit)sender;
            string formatValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewShreder.SetFocusedRowCellValue(GridColShrderEtcT, formatValue);
        }
        private void RepoTxtShrederTime1_EditValueChanged(object sender, EventArgs e)//
        {
            TextEdit textEdit = (TextEdit)sender;
            string formatValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewShreder.SetFocusedRowCellValue(GridColShrderEndTime, formatValue);
        }
        private void RepoTxtShrederTime2_EditValueChanged(object sender, EventArgs e)//
        {
            TextEdit textEdit = (TextEdit)sender;
            string formatValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewShreder.SetFocusedRowCellValue(GridColShrderStartTime, formatValue);
        }

        private void RepoTxtShrederTimeUSING_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;
            string formatValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewShreder.SetFocusedRowCellValue(GridColShrderOpTime, formatValue);
        }
        //


        private void RepoGilloTxtTime_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;
            string GilloValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewGuillotine.SetFocusedRowCellValue(GridColGilloStartTime, GilloValue);
        }

        private void RepoGilloTxtTimeEND_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;
            string GilloValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewGuillotine.SetFocusedRowCellValue(GridColGilloEndTime, GilloValue);
        }
        //


        private void RepoTxtTime_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;
            string WorkValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewWork.SetFocusedRowCellValue(GridColWorkStrtTime, WorkValue);
        }

        private void RepoTxtTimeEND_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;
            string WorkValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewWork.SetFocusedRowCellValue(GridColWorkEndTime, WorkValue);
        }

        //
        private void RepoTxtMntncTime_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;
            string MtncValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewMaintenance.SetFocusedRowCellValue(GridColMtncStartTime, MtncValue);
        }

        private void RepoTxtMntncTimeEND_EditValueChanged(object sender, EventArgs e)
        {
            TextEdit textEdit = (TextEdit)sender;
            string MtncValue = Convert.ToDateTime(textEdit.EditValue).ToString("HH:mm");
            GridViewMaintenance.SetFocusedRowCellValue(GridColMtncEndTime, MtncValue);
        }

        private DataTable SetM5Device(string sEQUIPCD)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" SELECT A.MG_NO                   ");
            strSql.AppendLine("      ,A.EQUIP_NM                    ");
            strSql.AppendLine("   FROM equip_cd A                 ");
            strSql.AppendLine("  WHERE A.MG_NO ='" + sEQUIPCD + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }
        private DataTable SetM5CHRG(string sCHRGCD)
        {
            StringBuilder strSql = new StringBuilder();
            strSql.Clear();
            strSql.AppendLine(" SELECT A.EMP_ID                   ");
            strSql.AppendLine("      ,A.EMP_NM                    ");
            strSql.AppendLine("   FROM hr_emp_basis A                 ");
            strSql.AppendLine("  WHERE A.EMP_ID ='" + sCHRGCD + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }
    }
}