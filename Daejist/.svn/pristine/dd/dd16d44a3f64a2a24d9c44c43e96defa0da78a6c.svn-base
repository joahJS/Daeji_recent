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
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.BandedGrid.ViewInfo;

namespace AccAdm
{
    public partial class GE003F00 : DevExpress.XtraEditors.XtraForm
    {
        public GE003F00()
        {
            InitializeComponent();
        }

        public string _EMPID = string.Empty;
        public string _BASDT = string.Empty;

        private void GE003F00_Load(object sender, EventArgs e)
        {
            SetLoadFormLayout();
            ComnEtcFunc.gp_SetColorFocused(layoutControl2);
            ComnEtcFunc.SetDateToValue(DateYY);
        }

        #region 초기 Layout, 권한 세팅
        public GridView[] arrGrdView;
        public DataRow rowUserInfo { get; set; }
        private void SetLoadFormLayout()
        {
            arrGrdView = new GridView[] { BGridViewRetr1, GridViewRetr2 };
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            NewYnchaData(); //조회시 연차생성
            SetGridViewRetr1Data();
        }

        //조회시 연차생성
        private void NewYnchaData()
        {
            try
            {
                string sYY = DateYY.EditValue?.ToString().Substring(0, 4);
                double dYncnt = 25;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" SELECT EMP_ID                     ");
                strSql.AppendLine("      , '" + sYY + "-01-01' AS BASDT");
                strSql.AppendLine("   FROM HR_EMP_BASIS               ");
                strSql.AppendLine("  WHERE EMPL_GB = 'Y'              ");
                strSql.AppendLine("    AND GDAYYN = 'Y'               ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dt != null)
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sEmpid = dt.Rows[i]["EMP_ID"]?.ToString();
                        string sBasdt = dt.Rows[i]["BASDT"]?.ToString();
                        string sUser = FmMainToolBar2.drUser["USRCD"]?.ToString();

                        strSql.Clear();
                        strSql.AppendLine(" IF EXISTS(SELECT EMPID FROM YNCH_M WHERE EMPID = '" + sEmpid + "' AND BASDT = '" + sBasdt + "') ");
                        strSql.AppendLine("    BEGIN                                                            ");
                        strSql.AppendLine("            SELECT '' AS MSG ");
                        strSql.AppendLine("      END                                                            ");
                        strSql.AppendLine(" ELSE                                                                ");
                        strSql.AppendLine("     BEGIN                                                           ");
                        strSql.AppendLine("           INSERT INTO YNCH_M(EMPID, BASDT, YNCNT, YNGUB, CUSER)            ");
                        strSql.AppendLine("                       VALUES('" + sEmpid + "', '" + sBasdt + "', " + dYncnt + ", '0', '" + sUser + "')                         ");
                        strSql.AppendLine("       END                                                           ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                }
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        public void SetGridViewRetr1Data()
        {
            try
            {
                string sDate = DateYY.EditValue?.ToString().Substring(0, 4);

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" WITH TEMP1 AS(                                                               ");
                strSql.AppendLine("     --총량 연차, 휴가                                                        ");
                strSql.AppendLine("     SELECT A1.EMP_ID                                                         ");
                strSql.AppendLine("          , A1.EMP_NM                                                         ");
                strSql.AppendLine("          , A1.DEPT_CD  ");
                strSql.AppendLine("          , A1.GRADE_CD ");
                strSql.AppendLine("          , B1.DEPT_NM                                                        ");
                strSql.AppendLine("          , CONVERT(DATE, A1.ENTRANCE_YMD) AS SRTDT                           ");
                strSql.AppendLine("          , SUM(CASE WHEN A2.YNGUB = '0' THEN A2.YNCNT ELSE 0 END) AS YNCNT1  ");
                strSql.AppendLine("          , SUM(CASE WHEN A2.YNGUB = '1' THEN A2.YNCNT ELSE 0 END) AS YNCNT2  ");
                strSql.AppendLine("       FROM HR_EMP_BASIS A1                                                   ");
                strSql.AppendLine("       LEFT JOIN YNCH_M A2                                                    ");
                strSql.AppendLine("         ON A1.EMP_ID = A2.EMPID                                              ");
                strSql.AppendLine("        AND SUBSTRING(A2.BASDT, 1, 4) = '"+ sDate + "'                                ");
                strSql.AppendLine("       LEFT JOIN ACC_DEPT_CD B1                                               ");
                strSql.AppendLine("         ON A1.DEPT_CD = B1.DEPT_CD                                           ");
                strSql.AppendLine("      WHERE A1.EMPL_GB = 'Y'                                                  ");
                strSql.AppendLine("        AND A1.GDAYYN = 'Y'                                                   ");
                strSql.AppendLine("      GROUP BY A1.EMP_ID, A1.EMP_NM, A1.DEPT_CD, A1.GRADE_CD, B1.DEPT_NM, A1.ENTRANCE_YMD    ");
                strSql.AppendLine(" ), TEMP2 AS(                                                                 ");
                strSql.AppendLine("     SELECT EMP_ID                                                            ");
                strSql.AppendLine("          , [연차01], [연차02], [연차03], [연차04], [연차05], [연차06]        ");
                strSql.AppendLine("          , [연차07], [연차08], [연차09], [연차10], [연차11], [연차12]        ");
                strSql.AppendLine("          , [휴가01], [휴가02], [휴가03], [휴가04], [휴가05], [휴가06]        ");
                strSql.AppendLine("          , [휴가07], [휴가08], [휴가09], [휴가10], [휴가11], [휴가12]        ");
                strSql.AppendLine("        FROM (--사용 연차(공용)                                               ");
                strSql.AppendLine("                 SELECT A1.EMP_ID                                             ");
				strSql.AppendLine("                      , '연차' + SUBSTRING(A2.BASDT, 6, 2) AS PART            ");
                strSql.AppendLine("                      , COUNT(A2.EMPID) AS USCNT                              ");
                strSql.AppendLine("                   FROM HR_EMP_BASIS A1                                       ");
                strSql.AppendLine("                   LEFT JOIN GDAYF A2                                         ");
                strSql.AppendLine("                     ON A1.EMP_ID = A2.EMPID                                  ");
                strSql.AppendLine("                    AND SUBSTRING(A2.BASDT,1,4) = '"+ sDate + "'              ");
                strSql.AppendLine("                    AND A2.RK1 LIKE '%연차(공용)%'");
                strSql.AppendLine("                  WHERE A1.EMPL_GB = 'Y'                                      ");
                strSql.AppendLine("                    AND A1.GDAYYN = 'Y'                                       ");
                strSql.AppendLine("                  GROUP BY A1.EMP_ID, SUBSTRING(A2.BASDT, 6, 2)               ");
                strSql.AppendLine("                  UNION ALL                                                   ");
                strSql.AppendLine("                 --사용 특별휴가.연차(개인), 병가, 반차                       ");
                strSql.AppendLine("                 SELECT A1.EMP_ID                                             ");
				strSql.AppendLine("                      , '휴가' + SUBSTRING(A2.BASDT, 6, 2) AS PART            ");
                strSql.AppendLine("                      , SUM(A2.GWKTM1) AS USCNT");
                //strSql.AppendLine("                      , SUM(CASE WHEN A2.SRTGB IN('YG', 'BG') THEN 1          ");
                //strSql.AppendLine("                                 WHEN A2.SRTGB = 'YA' OR A2.ENDGB = 'YP' THEN 0.5 ELSE 0 END) AS USCNT");
                strSql.AppendLine("                   FROM HR_EMP_BASIS A1                                                               ");
                strSql.AppendLine("                   LEFT JOIN GDAYF A2                                                                 ");
                strSql.AppendLine("                     ON A1.EMP_ID = A2.EMPID                                                          ");
                strSql.AppendLine("                    AND SUBSTRING(A2.BASDT,1,4) = '"+ sDate + "'                                              ");
                strSql.AppendLine("                    AND A2.GWKTM1 > 0");
                //strSql.AppendLine("                    AND(A2.SRTGB IN('YA', 'YG', 'BG') OR A2.ENDGB IN('YP'))                           ");
                strSql.AppendLine("                  WHERE A1.EMPL_GB = 'Y'                                                              ");
                strSql.AppendLine("                    AND A1.GDAYYN = 'Y'                                                               ");
                strSql.AppendLine("                  GROUP BY A1.EMP_ID, SUBSTRING(A2.BASDT, 6, 2)) AS TEMP                              ");
                strSql.AppendLine("        PIVOT(                                                                                        ");
                strSql.AppendLine("             MIN(USCNT) FOR PART IN( [연차01], [연차02], [연차03], [연차04], [연차05], [연차06]       ");
                strSql.AppendLine("                                 , [연차07], [연차08], [연차09], [연차10], [연차11], [연차12]         ");
                strSql.AppendLine("                                 , [휴가01], [휴가02], [휴가03], [휴가04], [휴가05], [휴가06]         ");
                strSql.AppendLine("                                 , [휴가07], [휴가08], [휴가09], [휴가10], [휴가11], [휴가12])        ");
	            strSql.AppendLine("        ) AS PVT                                                                                      ");
                strSql.AppendLine(" )                                                                                                    ");
                strSql.AppendLine(" SELECT 'N' AS CHK                                                                                          ");
                strSql.AppendLine("      , Z1.*                                                                                          ");
                strSql.AppendLine("      , Z1.YNCNT1 - Z1.USYN1 AS JAN1                                                                  ");
                strSql.AppendLine("      , Z1.YNCNT2 - Z1.USYN2 AS JAN2                                                                  ");
                strSql.AppendLine("   FROM(SELECT A1.EMP_ID                                                                              ");
                strSql.AppendLine("              , A1.EMP_NM                                                                             ");
                strSql.AppendLine("              , A1.DEPT_CD  ");
                strSql.AppendLine("              , A1.GRADE_CD ");
                strSql.AppendLine("              , A1.DEPT_NM                                                                            ");
                strSql.AppendLine("              , A1.SRTDT                                                                              ");
                strSql.AppendLine("              , A1.YNCNT1--총연차                                                                     ");
                strSql.AppendLine("              , ISNULL(A2.[연차01], 0) + ISNULL(A2.[연차02], 0) + ISNULL(A2.[연차03], 0) + ISNULL(A2.[연차04], 0) + ISNULL(A2.[연차05], 0) + ISNULL(A2.[연차06], 0)                       ");
                strSql.AppendLine("                + ISNULL(A2.[연차07], 0) + ISNULL(A2.[연차08], 0) + ISNULL(A2.[연차09], 0) + ISNULL(A2.[연차10], 0) + ISNULL(A2.[연차11], 0) + ISNULL(A2.[연차12], 0) AS USYN1--연차사용량");
                strSql.AppendLine("              , A1.YNCNT2--총휴가                                                                                                                                                         ");
                strSql.AppendLine("              , ISNULL(A2.[휴가01], 0) + ISNULL(A2.[휴가02], 0) + ISNULL(A2.[휴가03], 0) + ISNULL(A2.[휴가04], 0) + ISNULL(A2.[휴가05], 0) + ISNULL(A2.[휴가06], 0)                       ");
                strSql.AppendLine("                + ISNULL(A2.[휴가07], 0) + ISNULL(A2.[휴가08], 0) + ISNULL(A2.[휴가09], 0) + ISNULL(A2.[휴가10], 0) + ISNULL(A2.[휴가11], 0) + ISNULL(A2.[휴가12], 0) AS USYN2--휴가사용량");
                strSql.AppendLine("              , A2.[연차01], A2.[연차02], A2.[연차03], A2.[연차04], A2.[연차05], A2.[연차06] ");
                strSql.AppendLine("              , A2.[연차07], A2.[연차08], A2.[연차09], A2.[연차10], A2.[연차11], A2.[연차12] ");
                strSql.AppendLine("              , A2.[휴가01], A2.[휴가02], A2.[휴가03], A2.[휴가04], A2.[휴가05], A2.[휴가06] ");
                strSql.AppendLine("              , A2.[휴가07], A2.[휴가08], A2.[휴가09], A2.[휴가10], A2.[휴가11], A2.[휴가12] ");
                strSql.AppendLine("           FROM TEMP1 A1                                                                     ");
                strSql.AppendLine("           LEFT JOIN TEMP2 A2                                                                ");
                strSql.AppendLine("             ON A1.EMP_ID = A2.EMP_ID)Z1                                                     ");
                strSql.AppendLine("  ORDER BY Z1.DEPT_CD, Z1.GRADE_CD, Z1.EMP_NM                                                                         ");

                DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                GridRetr1.DataSource = dt;

                if (!string.IsNullOrEmpty(_EMPID))
                {
                    BGridViewRetr1.ClearSelection();
                    BGridViewRetr1.FocusedRowHandle = BGridViewRetr1.LocateByDisplayText(0, GridColEMPID, _EMPID);
                    if (BGridViewRetr1.FocusedRowHandle == 0)
                    {
                        SetGridRetr2Data();
                    }
                    _EMPID = string.Empty;
                }
                else
                {
                    if (BGridViewRetr1.FocusedRowHandle == 0)
                    {
                        SetGridRetr2Data();
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BGridViewRetr1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            SetGridRetr2Data();
        }

        public void SetGridRetr2Data()
        {
            try
            {
                string sEmpid = BGridViewRetr1.GetFocusedRowCellValue("EMP_ID")?.ToString();
                string sYY = DateYY.EditValue?.ToString().Substring(0, 4);

                if (!string.IsNullOrEmpty(sEmpid))
                {
                    StringBuilder strSql = new StringBuilder();

                    strSql.Clear();
                    strSql.AppendLine(" SELECT A1.EMPID                                                                                                ");
                    strSql.AppendLine("      , A2.EMP_NM                                                                                               ");
                    strSql.AppendLine("      , A1.BASDT                                                                                                ");
                    strSql.AppendLine("      , CASE WHEN A1.YNGUB = '0' THEN '연차' WHEN A1.YNGUB = '1' THEN '특별휴가' END AS YNGUB_1                 ");
                    strSql.AppendLine("      , A1.YNGUB                                                                                                ");
                    strSql.AppendLine("      , A1.YNCNT                                                                                                ");
                    strSql.AppendLine("      , A1.RK                                                                                                   ");
                    strSql.AppendLine("      , A1.CDATE                                                                                                ");
                    strSql.AppendLine("      , CASE WHEN TRY_PARSE(A1.CUSER AS NUMERIC) IS NULL THEN A1.CUSER ELSE DBO.FN_USRNM(A1.CUSER) END AS CUSER ");
                    strSql.AppendLine("      , A1.MDATE                                                                                                ");
                    strSql.AppendLine("      , CASE WHEN TRY_PARSE(A1.MUSER AS NUMERIC) IS NULL THEN A1.MUSER ELSE DBO.FN_USRNM(A1.MUSER) END AS MUSER ");
                    strSql.AppendLine("   FROM YNCH_M A1                                                                                               ");
                    strSql.AppendLine("   LEFT JOIN HR_EMP_BASIS A2                                                                                    ");
                    strSql.AppendLine("     ON A1.EMPID = A2.EMP_ID                                                                                    ");
                    strSql.AppendLine("  WHERE A1.EMPID = '" + sEmpid + "'                                                                                ");
                    strSql.AppendLine("    AND A1.BASDT BETWEEN '"+ sYY + "-01-01' AND '"+ sYY + "-12-30'");

                    DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    GridRetr2.DataSource = dt;

                    if (!string.IsNullOrEmpty(_BASDT))
                    {
                        GridViewRetr2.FocusedRowHandle = GridViewRetr2.LocateByDisplayText(0, GridColBASDT, _BASDT);
                        _BASDT = string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void GridViewRetr2_RowClick(object sender, RowClickEventArgs e)
        {
            if (e.Clicks == 2)
            {
                string sEmpid = GridViewRetr2.GetFocusedRowCellValue("EMPID")?.ToString();

                GE003F01 frm = new GE003F01();

                frm.Owner = this;
                frm.AddModifyGb = "MOD";
                frm._EMPID = sEmpid;
                frm._Row = GridViewRetr2.GetFocusedDataRow();
                if (frm.ShowDialog() == DialogResult.OK)
                {

                }
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string sEmpid = GridViewRetr2.GetFocusedRowCellValue("EMPID")?.ToString();

            GE003F01 frm = new GE003F01();

            frm.Owner = this;
            frm.AddModifyGb = "ADD";
            frm._EMPID = sEmpid;
            if (frm.ShowDialog() == DialogResult.OK)
            {

            }
        }

        private void BtnAddAll_Click(object sender, EventArgs e)
        {
            int[] selRows = BGridViewRetr1.GetSelectedRows();

            if (selRows.Length == 0)
            {
                XtraMessageBox.Show("연차 일괄생성할 데이터를 선택해주세요.");
                return;
            }

            GE003F01 frm = new GE003F01();

            frm.Owner = this;
            frm.AddModifyGb = "ALL";
            frm.DataRowSendEvent += new GE003F01.SendDataHandler(GetDataRow);
            if (frm.ShowDialog() == DialogResult.OK)
            {

            }
        }

        public void GetDataRow(Dictionary<string, string> dicParams)
        {
            //int[] selRows = BGridViewRetr1.GetSelectedRows();

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                string sBASDT = dicParams["BASDT"]?.ToString();
                string sYNCNT = dicParams["YNCNT"]?.ToString();
                string sYNGUB = dicParams["YNGUB"]?.ToString();
                string sRK = dicParams["RK"]?.ToString();
                string sUSER = FmMainToolBar2.drUser["USRCD"]?.ToString();

                double dYNCNT = 0;

                double.TryParse(sYNCNT, out dYNCNT);

                //for (int i = 0; i < selRows.Length; i++)
                //{
                //    string sEMPID = BGridViewRetr1.GetRowCellValue(selRows[i], "EMP_ID")?.ToString();

                //    strSql.Clear();
                //    strSql.AppendLine(" IF EXISTS(SELECT * FROM YNCH_M WHERE EMPID = '" + sEMPID + "' AND BASDT = '" + sBASDT + "')");
                //    strSql.AppendLine("     BEGIN                                                     ");
                //    strSql.AppendLine("             UPDATE YNCH_M                                     ");
                //    strSql.AppendLine("                SET YNCNT = " + dYNCNT + "                                  ");
                //    strSql.AppendLine("                  , YNGUB = '" + sYNGUB + "'");
                //    strSql.AppendLine("                  , RK = '" + sRK + "'                                    ");
                //    strSql.AppendLine("                  , MUSER = " + sUSER + "                                  ");
                //    strSql.AppendLine("                  , MDATE = CONVERT(VARCHAR(20), GETDATE(), 20)");
                //    strSql.AppendLine("             WHERE EMPID = '" + sEMPID + "' AND BASDT = '" + sBASDT + "'                   ");
                //    strSql.AppendLine("         END                                                   ");
                //    strSql.AppendLine(" ELSE                                                          ");
                //    strSql.AppendLine("     BEGIN                                                     ");
                //    strSql.AppendLine("             INSERT INTO YNCH_M( EMPID                         ");
                //    strSql.AppendLine("                               , BASDT                         ");
                //    strSql.AppendLine("                               , YNGUB");
                //    strSql.AppendLine("                               , YNCNT                         ");
                //    strSql.AppendLine("                               , RK                            ");
                //    strSql.AppendLine("                               , CUSER )                       ");
                //    strSql.AppendLine("                         VALUES( '" + sEMPID + "'                        ");
                //    strSql.AppendLine("                               , '" + sBASDT + "'                       ");
                //    strSql.AppendLine("                               , '" + sYNGUB + "'");
                //    strSql.AppendLine("                               , " + dYNCNT + "                         ");
                //    strSql.AppendLine("                               , '" + sRK + "'                            ");
                //    strSql.AppendLine("                               , " + sUSER + ")                        ");
                //    strSql.AppendLine("         END                                                   ");

                //    cmd.CommandType = CommandType.Text;
                //    cmd.CommandText = strSql.ToString();
                //    cmd.ExecuteNonQuery();
                //}

                DataTable dt = GridRetr1.DataSource as DataTable;

                if(dt != null)
                {
                    foreach(DataRow row in dt.Rows)
                    {
                        string sEMPID = row["EMP_ID"]?.ToString();
                        string sCHK = row["CHK"]?.ToString();

                        if (sCHK.Equals("Y"))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" IF EXISTS(SELECT * FROM YNCH_M WHERE EMPID = '" + sEMPID + "' AND BASDT = '" + sBASDT + "')");
                            strSql.AppendLine("     BEGIN                                                     ");
                            strSql.AppendLine("             UPDATE YNCH_M                                     ");
                            strSql.AppendLine("                SET YNCNT = " + dYNCNT + "                                  ");
                            strSql.AppendLine("                  , YNGUB = '" + sYNGUB + "'");
                            strSql.AppendLine("                  , RK = '" + sRK + "'                                    ");
                            strSql.AppendLine("                  , MUSER = " + sUSER + "                                  ");
                            strSql.AppendLine("                  , MDATE = CONVERT(VARCHAR(20), GETDATE(), 20)");
                            strSql.AppendLine("             WHERE EMPID = '" + sEMPID + "' AND BASDT = '" + sBASDT + "'                   ");
                            strSql.AppendLine("         END                                                   ");
                            strSql.AppendLine(" ELSE                                                          ");
                            strSql.AppendLine("     BEGIN                                                     ");
                            strSql.AppendLine("             INSERT INTO YNCH_M( EMPID                         ");
                            strSql.AppendLine("                               , BASDT                         ");
                            strSql.AppendLine("                               , YNGUB");
                            strSql.AppendLine("                               , YNCNT                         ");
                            strSql.AppendLine("                               , RK                            ");
                            strSql.AppendLine("                               , CUSER )                       ");
                            strSql.AppendLine("                         VALUES( '" + sEMPID + "'                        ");
                            strSql.AppendLine("                               , '" + sBASDT + "'                       ");
                            strSql.AppendLine("                               , '" + sYNGUB + "'");
                            strSql.AppendLine("                               , " + dYNCNT + "                         ");
                            strSql.AppendLine("                               , '" + sRK + "'                            ");
                            strSql.AppendLine("                               , " + sUSER + ")                        ");
                            strSql.AppendLine("         END                                                   ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();
                            cmd.ExecuteNonQuery();
                        }
                    }
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("일괄등록을 완료했습니다.");

                BtnRetr.PerformClick();
                //BGridViewRetr1.FocusedRowHandle = selRows[0];
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnDel_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
                {
                    XtraMessageBox.Show("해당 사용자에 대하여 삭제 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                    return;
                }

                string sEmpid = GridViewRetr2.GetFocusedRowCellValue("EMPID")?.ToString();
                string sEmpNm = GridViewRetr2.GetFocusedRowCellValue("EMP_NM")?.ToString();
                string sBasdt = GridViewRetr2.GetFocusedRowCellValue("BASDT")?.ToString();

                int idx1 = BGridViewRetr1.FocusedRowHandle;
                int idx = GridViewRetr2.FocusedRowHandle;
                _EMPID = sEmpid;

                if (string.IsNullOrEmpty(sEmpid) || string.IsNullOrEmpty(sBasdt))
                {
                    XtraMessageBox.Show("삭제할 항목이 없습니다.");
                    return;
                }

                if (XtraMessageBox.Show("발생일자 : " + sBasdt +
                    "\r\n 이름: "+ sEmpNm +
                     "\r\n선택된 항목을 삭제하시겠습니까? \r\n삭제한 데이터는 복구할 수 없습니다."
                   , "연차발생삭제 확인", MessageBoxButtons.YesNo) != DialogResult.Yes)
                {
                    return;
                }

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                StringBuilder strSql = new StringBuilder();

                strSql.Clear();
                strSql.AppendLine(" DELETE FROM YNCH_M");
                strSql.AppendLine("  WHERE EMPID = '"+ sEmpid + "' ");
                strSql.AppendLine("    AND BASDT = '"+ sBasdt + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                XtraMessageBox.Show("삭제가 완료되었습니다.");

                BtnRetr.PerformClick();
                BGridViewRetr1.FocusInvalidRow();
                BGridViewRetr1.FocusedRowHandle = idx1;
                GridViewRetr2.FocusedRowHandle = idx - 1;
            }
            catch(Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BGridViewRetr1_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void BGridViewRetr1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            ComnGridFunc.SettingGridViewRowIndicator(sender, e);
        }

        private void GE003F00_TextChanged(object sender, EventArgs e)
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

        private void GE003F00_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5)
                BtnRetr.PerformClick();
        }

        private void DateYY_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnPrev_Click(object sender, EventArgs e)
        {
            string sDate = DateYY.EditValue?.ToString();
            string sPrevDate = ComnEtcFunc.PrevYear(sDate);

            if (!string.IsNullOrEmpty(sPrevDate))
            {
                DateYY.EditValue = sPrevDate;

                BtnRetr.PerformClick();
            }
        }

        private void BtnNext_Click(object sender, EventArgs e)
        {
            string sDate = DateYY.EditValue?.ToString();
            string sNextDate = ComnEtcFunc.NextYear(sDate);

            if (!string.IsNullOrEmpty(sNextDate))
            {
                DateYY.EditValue = sNextDate;

                BtnRetr.PerformClick();
            }
        }

        private void BGridViewRetr1_MouseDown(object sender, MouseEventArgs e)
        {
            BandedGridView view = (BandedGridView)sender;
            BandedGridHitInfo hitInfo = view.CalcHitInfo(e.Location);

            //전체선택
            if (hitInfo.InBandPanel)
            {
                if (hitInfo.Band == null)
                    return;

                if (hitInfo.Band == gridBand1)
                {
                    string sYn = string.Empty;
                    for (int i = 0; i < view.RowCount; i++)
                    {
                        if (i == 0)
                        {
                            sYn = view.GetRowCellValue(i, GridColChk)?.ToString();
                        }

                        if (sYn.Equals("Y"))
                        {
                            view.SetRowCellValue(i, GridColChk, "N");
                        }
                        else
                        {
                            view.SetRowCellValue(i, GridColChk, "Y");
                        }
                    }

                    if (sYn.Equals("Y"))
                    {
                        view.ClearSelection();
                    }
                    else
                    {
                        view.SelectAll();
                    }
                }
            }
            else if (hitInfo.InColumnPanel)
            {
                if (hitInfo.Column == null)
                    return;

                if (hitInfo.Column == GridColChk)
                {
                    string sYn = string.Empty;
                    for (int i = 0; i < view.RowCount; i++)
                    {
                        if (i == 0)
                        {
                            sYn = view.GetRowCellValue(i, GridColChk)?.ToString();
                        }

                        if (sYn.Equals("Y"))
                        {
                            view.SetRowCellValue(i, GridColChk, "N");
                        }
                        else
                        {
                            view.SetRowCellValue(i, GridColChk, "Y");
                        }
                    }

                    if (sYn.Equals("Y"))
                    {
                        view.ClearSelection();
                    }
                    else
                    {
                        view.SelectAll();
                    }
                }
            }
        }

        private void repositoryItemCheckEdit1_CheckedChanged(object sender, EventArgs e)
        {
            int iIdx = BGridViewRetr1.FocusedRowHandle;
            string sChk = BGridViewRetr1.GetRowCellValue(iIdx, "CHK")?.ToString();

            if (string.IsNullOrEmpty(sChk))
                return;

            BGridViewRetr1.Focus();

            if (sChk.Equals("Y"))
            {
                BGridViewRetr1.SetRowCellValue(iIdx, "CHK", "N");
                BGridViewRetr1.UnselectRow(iIdx);
            }
            else
            {
                BGridViewRetr1.SetRowCellValue(iIdx, "CHK", "Y");
                BGridViewRetr1.SelectRow(iIdx);
            }
        }
    }
}