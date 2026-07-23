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
using DevExpress.XtraEditors.Repository;
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
*            2. 레이아웃 전체 저장 설정
*            
* 수정일자 : 2023-01-05
* 수정자   : 정은영
* 수정내용 : (현업요청)
*            1. 단가비고 입력 첫번째 탭으로 이동. 스크랩 상단, 슈레더 하단
*            2. 이력조회 두번째 탭으로 이동.
*/
namespace AccAdm
{
    public partial class UnitCostMgt : DevExpress.XtraEditors.XtraForm
    {
        public UnitCostMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void UnitCostMgt_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            ComnGridFunc.GridStyleBasicSetting(GridViewRetr2);

            ////////
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "AccDealerCdDev");

            DateEditFrom.EditValue = DateTime.Now.ToString("yyyy") + "-" + "01" + "-" + "01";
            DateEditTo.EditValue = DateTime.Now;

            DataTable dtUserCd = GetLookUpData("2", "", "", "");
            RepositoryItemGridLookUpEdit userLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(userLkup, dtUserCd, GridRetr, GridColUserId, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(userLkup, dtUserCd, GridLow, GridColLowUserCd, "CD", "NM", "");

            BtnRetr.PerformClick();

            arrGrdView = new GridView[] { GridViewRetr, GridViewLow };
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
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam, string sOther)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("1"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            else if (sNullYn.Equals("2"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine(" UNION ALL");
            }


            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'COM_COM_CD'");

            }
            if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.USRCD AS CD");
                strSql.AppendLine("      , A.USRNM AS NM");
                strSql.AppendLine("      , A.USRCD AS SEQ");
                strSql.AppendLine("   FROM ZUSRLST A");
                strSql.AppendLine("  WHERE 1=1");

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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if(xtraTabControl1.SelectedTabPageIndex == 0)
            {
                GetGridRetr();
                GetGridRetr2();
                GetGridRetr3();
            }
            else if (xtraTabControl1.SelectedTabPageIndex == 1)
            {
                GetGridRetr1();
                GetGridLowRetr();
            }
        }

        private void GetGridRetr()
        {
            string sGb = RdgbGb.EditValue.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("SELECT J_SPEC AS USERID"); 
            strSql.AppendLine("     , GOODS_ID");
            strSql.AppendLine("     , THICKPOS"); 
            strSql.AppendLine("     , DAEGUBUN ");
            strSql.AppendLine("     , GUBUN1");
            strSql.AppendLine("     , DANGA ");//인근(기준)
            strSql.AppendLine("     , SELLPRC1 ");//대한단가
            strSql.AppendLine("     , SELLPRC2 ");//현대제철
            strSql.AppendLine("     , SELLPRC3");//동국(포항)
            strSql.AppendLine("     , SELLPRC4 ");//포스코(광양)
            strSql.AppendLine("     , SELLPRC5 ");//포스코(포항)
            strSql.AppendLine("     , SELLPRC6 ");//한국철강
            strSql.AppendLine("     , SELLPRC7 ");//한국특강
            strSql.AppendLine("     , SELLPRC8 ");//세아창원
            strSql.AppendLine("     , SELLPRC9 ");
            strSql.AppendLine("     , LASTDATE "); 
            strSql.AppendLine("     , J_SERIAL "); 
            strSql.AppendLine("  FROM JAJAE");
            if (sGb.Equals("999"))strSql.AppendLine(" WHERE THICKPOS = '"+sGb+"'");
            else strSql.AppendLine(" WHERE THICKPOS <= '" + sGb + "'");
            strSql.AppendLine(" ORDER BY THICKPOS");

           DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

           GridRetr.DataSource = dt;
        }

        private void GetGridRetr1()
        {
            string sGb = RdgbGb.EditValue.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine("SELECT GUBUN1");
            strSql.AppendLine("     , J_SERIAL ");
            strSql.AppendLine("     , DAEGUBUN ");
            strSql.AppendLine("  FROM JAJAE");
            if (sGb.Equals("999")) strSql.AppendLine(" WHERE THICKPOS = '" + sGb + "'");
            else strSql.AppendLine(" WHERE THICKPOS <= '" + sGb + "'");
            strSql.AppendLine(" ORDER BY THICKPOS");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            GridGubun.DataSource = dt;
        }

        private void GetGridRetr2()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT TOP 1         ");
            strSql.AppendLine("        J_DATE        ");
            strSql.AppendLine("      , RK1           ");
            strSql.AppendLine("      , RK2           ");
            strSql.AppendLine("      , RK3           ");
            strSql.AppendLine("      , RK4           ");
            strSql.AppendLine("      , RK5           ");
            strSql.AppendLine("      , RK6           ");
            strSql.AppendLine("      , RK7           ");
            strSql.AppendLine("      , RK8           ");
            strSql.AppendLine("      , RK9           ");
            strSql.AppendLine("      , HDDT          ");
            strSql.AppendLine("      , DHDT          ");
            strSql.AppendLine("      , DGDT          ");
            strSql.AppendLine("      , PGDT          ");
            strSql.AppendLine("      , PPDT          ");
            strSql.AppendLine("      , HCDT          ");
            strSql.AppendLine("      , HTDT          ");
            strSql.AppendLine("      , SADT          ");
            strSql.AppendLine("      , INDT          ");
            strSql.AppendLine("      , SCRDT         ");
            strSql.AppendLine("   FROM JAJAERK       ");
            strSql.AppendLine("  ORDER BY J_DATE DESC");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr2.DataSource = dt;
        }

        private void GetGridRetr3()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT TOP 1         ");
            strSql.AppendLine("        J_DATE        ");
            strSql.AppendLine("      , YADT1         ");
            strSql.AppendLine("      , YADT2         ");
            strSql.AppendLine("      , PAPDT         ");
            strSql.AppendLine("   FROM JAJAERK       ");
            strSql.AppendLine("  ORDER BY J_DATE DESC");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr3.DataSource = dt;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if(xtraTabControl1.SelectedTabPageIndex == 0)
            {
                SaveDanga();
                SaveRkDate();
            }
        }

        private void SaveDanga()
        {

            DataTable dt = (DataTable)GridRetr.DataSource;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_SAVE(dt);            
            DataTable dtMod = dsSave.Tables[1];
                    
            string sThickPos = string.Empty;
            string sGubun = string.Empty;
            string sGrade = string.Empty;
            string sDanga = string.Empty;
            string sSellPrc1 = string.Empty;
            string sSellPrc2 = string.Empty;
            string sSellPrc3 = string.Empty;
            string sSellPrc4 = string.Empty;
            string sSellPrc5 = string.Empty;
            string sSellPrc6 = string.Empty;
            string sSellPrc7 = string.Empty;
            string sSellPrc8 = string.Empty;
            string sSellPrc9 = string.Empty;
            string sLastDate = string.Empty;
            double sJSerial = 0;
            
            StringBuilder insertSql = new StringBuilder();
            StringBuilder updateSql = new StringBuilder();
            StringBuilder strSql = new StringBuilder();

            if (dtMod.Rows.Count > 0)  // modify
            {
                try
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    for (int j = 0; j < dtMod.Rows.Count; j++)
                    {
                        string sUserCd = FmMainToolBar2.dtUserAutInfo.Rows[0]["USRCD"].ToString();

                        sThickPos = Convert.ToString(dtMod.Rows[j]["THICKPOS"]);
                        sDanga = dtMod.Rows[j]["DANGA"]?.ToString(); ;
                        sSellPrc1 = dtMod.Rows[j]["SELLPRC1"]?.ToString();
                        sSellPrc2 = dtMod.Rows[j]["SELLPRC2"]?.ToString();
                        sSellPrc3 = dtMod.Rows[j]["SELLPRC3"]?.ToString();
                        sSellPrc4 = dtMod.Rows[j]["SELLPRC4"]?.ToString();
                        sSellPrc5 = dtMod.Rows[j]["SELLPRC5"]?.ToString();
                        sSellPrc6 = dtMod.Rows[j]["SELLPRC6"]?.ToString();
                        sSellPrc7 = dtMod.Rows[j]["SELLPRC7"]?.ToString();
                        sSellPrc8 = dtMod.Rows[j]["SELLPRC8"]?.ToString();
                        sSellPrc9 = dtMod.Rows[j]["SELLPRC9"]?.ToString();

                        double dDanga = 0;
                        double dSellPrc1 = 0;
                        double dSellPrc2 = 0;
                        double dSellPrc3 = 0;
                        double dSellPrc4 = 0;
                        double dSellPrc5 = 0;
                        double dSellPrc6 = 0;
                        double dSellPrc7 = 0;
                        double dSellPrc8 = 0;
                        double dSellPrc9 = 0;

                        double.TryParse(sDanga, out dDanga);
                        double.TryParse(sSellPrc1, out dSellPrc1);
                        double.TryParse(sSellPrc2, out dSellPrc2);
                        double.TryParse(sSellPrc3, out dSellPrc3);
                        double.TryParse(sSellPrc4, out dSellPrc4);
                        double.TryParse(sSellPrc5, out dSellPrc5);
                        double.TryParse(sSellPrc6, out dSellPrc6);
                        double.TryParse(sSellPrc7, out dSellPrc7);
                        double.TryParse(sSellPrc8, out dSellPrc8);
                        double.TryParse(sSellPrc9, out dSellPrc9);

                        sLastDate = dtMod.Rows[j]["LASTDATE"]?.ToString();
                        if (string.IsNullOrEmpty(sLastDate))
                        {
                            sLastDate = DateTime.Today.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            sLastDate = sLastDate.Substring(0, 10);
                        }

                        sJSerial = Convert.ToDouble(dtMod.Rows[j]["J_SERIAL"]);
                        dtMod.Rows[j]["SELLPRC2"].Equals("0");

                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT * ");
                        strSql.AppendLine("   FROM JAJAE A");
                        strSql.AppendLine("  WHERE J_SERIAL = '" + sJSerial + "' ");

                        DataTable dtt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                        insertSql.Clear();
                        insertSql.AppendLine(" ");
                        insertSql.AppendLine(" INSERT INTO JAJAEHIS ( ");
                        insertSql.AppendLine("                            USERID, J_SERIAL, J_DATE ");
                        insertSql.AppendLine("                          , DAEGUBUN, GUBUN, YK_DANGA ");
                        insertSql.AppendLine("                          , GOJUNG_DANGA, DANGA, EJEUN_DANGA ");
                        insertSql.AppendLine("                          , SP_DANGA, IN_DANGA, OUT_DANGA ");
                        insertSql.AppendLine("                          , SEHA_DANGA, THICKPOS ");
                        insertSql.AppendLine("                          , Danga2, Danga3, Danga4");
                        insertSql.AppendLine("                          , Danga5, Danga6, Danga7");
                        insertSql.AppendLine("                          , Danga8, Danga9         ");
                        insertSql.AppendLine("                          ) ");
                        insertSql.AppendLine("                     SELECT J_SPEC AS J_SPEC ");
                        insertSql.AppendLine("                          , J_SERIAL  AS J_SERIAL");
                        insertSql.AppendLine("                          , CONVERT(VARCHAR(20),getdate(),20)  AS J_DATE");
                        insertSql.AppendLine("                          , DAEGUBUN AS DAEGUBUN");
                        insertSql.AppendLine("                          , GUBUN1 AS GUBUN ");
                        insertSql.AppendLine("                          , ISNULL(SELLPRC1,0) AS YK_DANGA       ");
                        insertSql.AppendLine("                          , ISNULL(SELLPRC2, 0) AS GOJUNG_DANGA  ");
                        insertSql.AppendLine("                          , ISNULL(DANGA, 0) AS DANGA            ");
                        insertSql.AppendLine("                          , ISNULL(SELLPRC3, 0) AS EJEUN_DANGA   ");
                        insertSql.AppendLine("                          , ISNULL(SELLPRC4, 0) AS SP_DANGA      ");
                        insertSql.AppendLine("                          , ISNULL(SELLPRC5, 0) AS IN_DANGA      ");
                        insertSql.AppendLine("                          , ISNULL(SELLPRC6, 0) AS OUT_DANGA     ");
                        insertSql.AppendLine("                          , ISNULL(SELLPRC7, 0) AS SEHA_DANGA    ");
                        insertSql.AppendLine("                          , THICKPOS ");
                        insertSql.AppendLine("                          , ISNULL(SellPrc2,0)");
                        insertSql.AppendLine("                          , ISNULL(SellPrc3,0)");
                        insertSql.AppendLine("                          , ISNULL(SellPrc4,0)");
                        insertSql.AppendLine("                          , ISNULL(sellprc5,0)");
                        insertSql.AppendLine("                          , ISNULL(sellprc6,0)");
                        insertSql.AppendLine("                          , ISNULL(sellprc7,0)");
                        insertSql.AppendLine("                          , ISNULL(sellprc8,0)");
                        insertSql.AppendLine("                          , ISNULL(sellprc9,0)");
                        insertSql.AppendLine("                       FROM  JAJAE ");
                        insertSql.AppendLine("                      WHERE  J_SERIAL = " + sJSerial + "");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = insertSql.ToString();
                        cmd.ExecuteNonQuery();

                        //table has no partition for value 202006 오류
                        //string sLogRmk = "Table:JAJAEHIS -> J_SERIAL:" + sJSerial;
                        //ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, (j + 1).ToString(), "I", this.Name, sLogRmk, cmd);

                        updateSql.Clear();
                        updateSql.AppendLine(" ");
                        updateSql.AppendLine(" UPDATE JAJAE");
                        updateSql.AppendLine("    SET THICKPOS = '" + sThickPos + "' ");
                        updateSql.AppendLine("      , DANGA = '" + dDanga + "' ");
                        updateSql.AppendLine("      , SELLPRC1 = '" + dSellPrc1 + "' ");
                        updateSql.AppendLine("      , SELLPRC2 = '" + dSellPrc2 + "' ");
                        updateSql.AppendLine("      , SELLPRC3 = '" + dSellPrc3 + "' ");
                        updateSql.AppendLine("      , SELLPRC4 = '" + dSellPrc4 + "' ");
                        updateSql.AppendLine("      , SELLPRC5 = '" + dSellPrc5 + "' ");
                        updateSql.AppendLine("      , SELLPRC6 = '" + dSellPrc6 + "' ");
                        updateSql.AppendLine("      , SELLPRC7 = '" + dSellPrc7 + "' ");
                        updateSql.AppendLine("      , SELLPRC8 = '" + dSellPrc8 + "' ");
                        updateSql.AppendLine("      , SELLPRC9 = '" + dSellPrc9 + "' ");
                        updateSql.AppendLine("      , LASTDATE = CONVERT(VARCHAR(10),getdate(),23)");
                        updateSql.AppendLine("      , J_SPEC = '" + sUserCd + "' ");
                        updateSql.AppendLine("  WHERE J_SERIAL = '" + sJSerial + "' ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = updateSql.ToString();
                        cmd.ExecuteNonQuery();

                        //string sLogRmk2 = "Table:JAJAE -> J_SERIAL:" + sJSerial;
                        //ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, (j + 1).ToString(), "U", this.Name, sLogRmk2, cmd);
                        
                    }
                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    GetGridRetr();

                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColJ_SERIAL, dtMod.Rows[0]["J_SERIAL"]?.ToString());
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                //XtraMessageBox.Show("변경된 내용이 없습니다.");
            }
        }

        private void SaveRkDate()
        {

            DataTable dt = (DataTable)GridRetr2.DataSource;
            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE(dt);
            DataTable dtMod = dsSave.Tables[0];

            DataTable dt2 = (DataTable)GridRetr3.DataSource;
            DataSet dsSave2 = ComGrid.GET_DATASET_FOR_MERGE(dt2);
            DataTable dtMod2 = dsSave2.Tables[0];

            string sRK1   = string.Empty;
            string sRK2   = string.Empty;
            string sRK3   = string.Empty;
            string sRK4   = string.Empty;
            string sRK5   = string.Empty;
            string sRK6   = string.Empty;
            string sRK7   = string.Empty;
            string sRK8   = string.Empty;
            string sRK9   = string.Empty;
            string sHDDT  = string.Empty;
            string sDHDT  = string.Empty;
            string sDGDT  = string.Empty;
            string sPGDT  = string.Empty;
            string sPPDT  = string.Empty;
            string sHCDT  = string.Empty;
            string sHTDT  = string.Empty;
            string sSADT  = string.Empty;
            string sINDT  = string.Empty;
            string sYADT1 = string.Empty;
            string sYADT2 = string.Empty;
            string sSCRDT = string.Empty;
            string sPAPDT = string.Empty;


            if (dtMod.Rows.Count > 0 || dtMod2.Rows.Count > 0)
            {
                try
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    StringBuilder strSql = new StringBuilder();

                    foreach(DataRow row in dt.Rows)
                    {
                        //스크랩
                        sRK1 = row["RK1"]?.ToString();
                        sRK2 = row["RK2"]?.ToString();
                        sRK3 = row["RK3"]?.ToString();
                        sRK4 = row["RK4"]?.ToString();
                        sRK5 = row["RK5"]?.ToString();
                        sRK6 = row["RK6"]?.ToString();
                        sRK7 = row["RK7"]?.ToString();
                        sRK8 = row["RK8"]?.ToString();
                        sRK9 = row["RK9"]?.ToString();
                        sHDDT = row["HDDT"]?.ToString();
                        sDHDT = row["DHDT"]?.ToString();
                        sDGDT = row["DGDT"]?.ToString();
                        sPGDT = row["PGDT"]?.ToString();
                        sPPDT = row["PPDT"]?.ToString();
                        sHCDT = row["HCDT"]?.ToString();
                        sHTDT = row["HTDT"]?.ToString();
                        sSADT = row["SADT"]?.ToString();
                        sINDT = row["INDT"]?.ToString();
                        sSCRDT = row["SCRDT"]?.ToString().Substring(0,10);
                    }

                    foreach(DataRow row in dt2.Rows)
                    {
                        //폐압
                        sYADT1 = row["YADT1"]?.ToString();
                        sYADT2 = row["YADT2"]?.ToString();
                        sPAPDT = row["PAPDT"]?.ToString().Substring(0, 10);
                    }

                    strSql.Clear();
                    strSql.AppendLine(" INSERT INTO JAJAERK(J_DATE                               ");
                    strSql.AppendLine("                    , RK1                                 ");
                    strSql.AppendLine("                    , RK2                                 ");
                    strSql.AppendLine("                    , RK3                                 ");
                    strSql.AppendLine("                    , RK4                                 ");
                    strSql.AppendLine("                    , RK5                                 ");
                    strSql.AppendLine("                    , RK6                                 ");
                    strSql.AppendLine("                    , RK7                                 ");
                    strSql.AppendLine("                    , RK8                                 ");
                    strSql.AppendLine("                    , RK9                                 ");
                    strSql.AppendLine("                    , HDDT                                ");
                    strSql.AppendLine("                    , DHDT                                ");
                    strSql.AppendLine("                    , DGDT                                ");
                    strSql.AppendLine("                    , PGDT                                ");
                    strSql.AppendLine("                    , PPDT                                ");
                    strSql.AppendLine("                    , HCDT                                ");
                    strSql.AppendLine("                    , HTDT                                ");
                    strSql.AppendLine("                    , SADT                                ");
                    strSql.AppendLine("                    , INDT                                ");
                    strSql.AppendLine("                    , YADT1                               ");
                    strSql.AppendLine("                    , YADT2                               ");
                    strSql.AppendLine("                    , SCRDT                               ");
                    strSql.AppendLine("                    , PAPDT)                              ");
                    strSql.AppendLine("                VALUES                                    ");
                    strSql.AppendLine("                     (CONVERT(VARCHAR(20), GETDATE(), 20) ");
                    strSql.AppendLine("                     , '" + sRK1 + "'                         ");
                    strSql.AppendLine("                     , '" + sRK2 + "'                         ");
                    strSql.AppendLine("                     , '" + sRK3 + "'                         ");
                    strSql.AppendLine("                     , '" + sRK4 + "'                         ");
                    strSql.AppendLine("                     , '" + sRK5 + "'                         ");
                    strSql.AppendLine("                     , '" + sRK6 + "'                         ");
                    strSql.AppendLine("                     , '" + sRK7 + "'                         ");
                    strSql.AppendLine("                     , '" + sRK8 + "'                         ");
                    strSql.AppendLine("                     , '" + sRK9 + "'                         ");
                    strSql.AppendLine("                     , '" + sHDDT + "'                        ");
                    strSql.AppendLine("                     , '" + sDHDT + "'                        ");
                    strSql.AppendLine("                     , '" + sDGDT + "'                        ");
                    strSql.AppendLine("                     , '" + sPGDT + "'                        ");
                    strSql.AppendLine("                     , '" + sPPDT + "'                        ");
                    strSql.AppendLine("                     , '" + sHCDT + "'                        ");
                    strSql.AppendLine("                     , '" + sHTDT + "'                        ");
                    strSql.AppendLine("                     , '" + sSADT + "'                        ");
                    strSql.AppendLine("                     , '" + sINDT + "'                        ");
                    strSql.AppendLine("                     , '" + sYADT1 + "'                       ");
                    strSql.AppendLine("                     , '" + sYADT2 + "'                       ");
                    strSql.AppendLine("                     , '" + sSCRDT + "'                       ");
                    strSql.AppendLine("                     , '" + sPAPDT + "')                  ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    MessageBox.Show("저장을 완료했습니다.");

                    GetGridRetr2();
                    GetGridRetr3();
                }
                catch (Exception ex)
                {
                    DBConn.dbTran.Rollback();
                    DBConn.dbTran = null;
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                //XtraMessageBox.Show("변경된 내용이 없습니다.");
            }
        }

        private void RdgbGb_EditValueChanged(object sender, EventArgs e)
        {
            GetGridRetr();
        }

        private void GetGridLowRetr()
        {
            GridLow.DataSource = null;

            string sDtFrom = DateEditFrom.EditValue?.ToString().Replace("-","").Substring(0,8);
            string sDtTo = DateEditTo.EditValue?.ToString().Replace("-", "").Substring(0, 8);
            string sResultYmdTo = string.Empty;

            if (!string.IsNullOrEmpty(sDtTo))
            {
                int dYear = Convert.ToInt32(sDtTo.Substring(0, 4));
                int dMonth = Convert.ToInt32(sDtTo.Substring(4, 2));
                int dDay = Convert.ToInt32(sDtTo.Substring(6, 2));

                DateTime date = new DateTime(dYear, dMonth, dDay);

                DateTime dtAdd = date.AddDays(+1);

                sResultYmdTo = dtAdd.ToString().Replace("-", "").Substring(0, 8);
            }

            string sDaeGb= GridViewGubun.GetFocusedRowCellValue("DAEGUBUN").ToString();
            string sGb = GridViewGubun.GetFocusedRowCellValue("GUBUN1").ToString();     
            
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT USERID ");
            strSql.AppendLine("      , J_SERIAL ");
            strSql.AppendLine("      , J_DATE ");
            strSql.AppendLine("      , DAEGUBUN ");
            strSql.AppendLine("      , GUBUN ");
            strSql.AppendLine("      , YK_DANGA ");
            strSql.AppendLine("      , GOJUNG_DANGA ");
            strSql.AppendLine("      , DANGA ");
            strSql.AppendLine("      , EJEUN_DANGA ");
            strSql.AppendLine("      , SP_DANGA ");
            strSql.AppendLine("      , IN_DANGA ");
            strSql.AppendLine("      , OUT_DANGA ");
            strSql.AppendLine("      , SEHA_DANGA ");
            strSql.AppendLine("      , THICKPOS ");
            strSql.AppendLine("      , Danga2 ");
            strSql.AppendLine("      , Danga3 ");
            strSql.AppendLine("      , Danga4 ");
            strSql.AppendLine("      , Danga5 ");
            strSql.AppendLine("      , Danga6 ");
            strSql.AppendLine("      , Danga7 ");
            strSql.AppendLine("      , Danga8 ");
            strSql.AppendLine("      , Danga9 ");
            strSql.AppendLine("   FROM JAJAEHIS ");
            strSql.AppendLine("  WHERE DAEGUBUN = '" + sDaeGb + "'");
            strSql.AppendLine("    AND  GUBUN = '" + sGb + "'");
            strSql.AppendLine("    AND  J_DATE >= '" + sDtFrom + "'");
            strSql.AppendLine("    AND  J_DATE <= '" + sResultYmdTo + "'");
            strSql.AppendLine("  ORDER BY J_DATE DESC ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridLow.DataSource = dt;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void UnitCostMgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { }
            else if (e.KeyCode == Keys.F1) { }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4) { }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if(e.KeyCode == Keys.F8)
            {
                BtnExcel.PerformClick();
            }
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void GridViewGubun_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (e.FocusedRowHandle >= 0)
            {
                GetGridLowRetr();
            }
            else if (e.FocusedRowHandle < 0)
            {
                GridLow.DataSource = null;
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
                string sFileNM = "단가 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;
                    if(xtraTabControl1.SelectedTabPageIndex == 0)
                    {
                        GridRetr.ExportToXls(FileName + ".xls");
                    }
                    else if (xtraTabControl1.SelectedTabPageIndex == 1)
                    {
                        GridLow.ExportToXls(FileName + ".xls");
                    }
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

        private void UnitCostMgt_TextChanged(object sender, EventArgs e)
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

        private void GridViewRetr_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr.UpdateCurrentRow();
        }

        private void RepoDateEdit_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if(DateTime.TryParse(sDate, out DateTime result))
            {
                GridViewRetr.SetRowCellValue(GridViewRetr.FocusedRowHandle, GridViewRetr.FocusedColumn, result.ToString("yyyy-MM-dd"));
            }
        }

        private void RepoRetr2DateEdit_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            string sDate = dateEdit.EditValue?.ToString();

            if(DateTime.TryParse(sDate, out DateTime result))
            {
                GridViewRetr2.SetRowCellValue(GridViewRetr2.FocusedRowHandle, GridViewRetr2.FocusedColumn, result.ToString("yyyy-MM-dd"));
            }
        }

        private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            BtnRetr.PerformClick();

            if(xtraTabControl1.SelectedTabPageIndex == 1)
            {
                layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Never;
            }
            else
            {
                layoutControlItem4.Visibility = DevExpress.XtraLayout.Utils.LayoutVisibility.Always;
            }
        }

        private void GridViewRetr2_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr2.UpdateCurrentRow();
        }

        private void GridViewRetr3_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewRetr3.UpdateCurrentRow();
        }

        private void DateEditTo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}