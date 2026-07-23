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
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraBars;
using DevExpress.XtraGrid.EditForm.Helpers.Controls;
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
    public partial class EquipConsumeCd : DevExpress.XtraEditors.XtraForm
    {
        public EquipConsumeCd()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void EquipConsumeMgt_Load(object sender, EventArgs e)
        {
            //
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DataTable dtMenuItem = GetLookUpData("1", "Y", "Y");
            DataTable dtDept = GetLookUpData("2", "N", "Y");

            ComLib.ComGrid.SetLookUpEdit(LkupMenuItemNm, dtMenuItem, "CD", "NM", "");
            LkupMenuItemNm.Properties.PopulateColumns();
            LkupMenuItemNm.Properties.Columns[0].Visible = false;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
            arrGrdView = new GridView[] { GridViewRetr };
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

            BtnRetr.PerformClick();
        }

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
            else
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '공용' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                //strSql.AppendLine(" SELECT A.CON_ITEM_CD AS CD");
                //strSql.AppendLine("      , A.CON_LARGE_CATE AS NM");
                //strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY CON_ITEM_NM) AS SEQ ");
                //strSql.AppendLine("   FROM EQUIP_CONSUME_MGT A");
                //strSql.AppendLine("  WHERE USE_YN = 'Y'");
                //strSql.AppendLine("  GROUP BY CON_LARGE_CATE, CON_ITEM_CD, CON_ITEM_NM");
                strSql.AppendLine(" SELECT A.CON_LARGE_CATE AS CD                                          ");
                strSql.AppendLine("      , A.CON_LARGE_CATE AS NM                            ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY CON_LARGE_CATE) AS SEQ ");
                strSql.AppendLine("   FROM EQUIP_CONSUME_MGT A                               ");
                strSql.AppendLine("  WHERE USE_YN = 'Y'                                      ");
                strSql.AppendLine("    AND CON_LARGE_CATE <> '' ");
                strSql.AppendLine("  GROUP BY CON_LARGE_CATE                                 ");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEPT_CD) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  WHERE USE_YN = 'Y'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE CD_GB = 'ITEM_INOUT_GB'");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT DEALER_CD AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
                strSql.AppendLine("  WHERE EOB_YN = 'N' ");
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            //DataTable dtChk = (DataTable)GridRetr.DataSource;
            //if (dtChk != null)
            //{
            //    DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridRetr.DataSource);
            //    DataTable dtMerge = dsSave.Tables[0];

            //    if (dtMerge.Rows.Count > 0)
            //    {
            //        if (XtraMessageBox.Show("리스트에 추가 및 수정사항이 있습니다. \r\n 그래도 조회하시겠습니까?"
            //        , "경고!", MessageBoxButtons.YesNo) == DialogResult.No)
            //        {
            //            return;
            //        }
            //    }
            //}

            string sItemNm = LkupMenuItemNm.Text;

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.CON_ITEM_CD ");
            strSql.AppendLine(" 	   , A.CON_ITEM_NM ");
            strSql.AppendLine(" 	   , A.CON_ITEM_SPEC ");
            strSql.AppendLine(" 	   , A.CON_ITEM_PURP ");
            strSql.AppendLine(" 	   , A.CON_LARGE_CATE ");
            strSql.AppendLine(" 	   , A.EXCHANGE_CYCLE ");
            strSql.AppendLine(" 	   , A.OPTIMAL_STOCK ");
            strSql.AppendLine(" 	   , A.ITEM_UNIT ");
            strSql.AppendLine(" 	   , A.STOCK_AMT ");
            strSql.AppendLine(" 	   , A.LACK_STOCK_AMT ");
            strSql.AppendLine(" 	   , A.NOTE ");
            strSql.AppendLine("        , A.USE_YN ");
            strSql.AppendLine(" 	   , A.ENT_DT ");
            strSql.AppendLine("        , CASE WHEN TRY_PARSE(A.ENT_ID  AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID");
	        strSql.AppendLine("        , A.MFY_DT                                                                                                 ");
            strSql.AppendLine("        , CASE WHEN TRY_PARSE(A.MFY_ID  AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID");
            strSql.AppendLine("   FROM EQUIP_CONSUME_MGT A ");
            strSql.AppendLine("  WHERE A.USE_YN = 'Y' ");
            if (!string.IsNullOrEmpty(sItemNm)) strSql.AppendLine("    AND A.CON_LARGE_CATE = '" + sItemNm + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            if (GridViewRetr.RowCount > 0)
            {
                string sCON_ITEM_NM = GridViewRetr.GetRowCellValue(GridViewRetr.RowCount - 1, "CON_ITEM_NM")?.ToString();

                if (string.IsNullOrEmpty(sCON_ITEM_NM))
                {
                    XtraMessageBox.Show("소모품명을 입력하세요.");
                    GridViewRetr.FocusedRowHandle = GridViewRetr.RowCount - 1;
                    GridViewRetr.FocusedColumn = GridColConItemCd;
                    GridViewRetr.Focus();
                    return;
                }
            }

            GridViewRetr.AddNewRow();
            GridViewRetr.Focus();
            GridViewRetr.SetFocusedRowCellValue("USE_YN", "Y");
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            DataSet dsSave = ComGrid.GET_DATASET_FOR_MERGE((DataTable)GridRetr.DataSource);
            DataTable dtMerge = dsSave.Tables[0];

            for (int i = 0; i < dtMerge.Rows.Count; i++)
            {
                if (String.IsNullOrEmpty(dtMerge.Rows[i]["CON_ITEM_NM"].ToString()))
                {
                    XtraMessageBox.Show("품명을 입력해주세요.");
                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0,GridColConItemNm, dtMerge.Rows[i]["CON_ITEM_NM"].ToString());
                    return;
                }
                if (String.IsNullOrEmpty(dtMerge.Rows[i]["USE_YN"].ToString()))
                {
                    XtraMessageBox.Show("사용구분을 입력해주세요.");
                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColUseYn, dtMerge.Rows[i]["USE_YN"].ToString());
                    return;
                }
            }

            string sConsumeCd = string.Empty;
            string sConItemNm = string.Empty;
            string sConItemSpec = string.Empty;
            string sConItemPurp = string.Empty;
            string sConLargeCate = string.Empty;
            string sExchangeCycle = string.Empty;
            double dOptimalStock = 0;
            string sItemUnit = string.Empty;
            double dStockAmt = 0;
            double dLackStockAmt = 0;
            string sNote = string.Empty;
            string sUseYn = string.Empty;
            string sId = FmMainToolBar2.drUser["USRCD"]?.ToString();

            try
            {

                if (dtMerge.Rows.Count > 0)
                {
                    DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                    SqlCommand cmd = DBConn.dbCon.CreateCommand();
                    cmd.Transaction = DBConn.dbTran;

                    for (int i = 0; i < dtMerge.Rows.Count; i++)
                    {
                        StringBuilder strSql = new StringBuilder();

                        sConsumeCd = dtMerge.Rows[i]["CON_ITEM_CD"].ToString();
                        if (String.IsNullOrEmpty(sConsumeCd))
                        {
                            strSql.Clear();
                            strSql.AppendLine(" SELECT MAX(A.CON_ITEM_CD) AS MAX_VALUE ");
                            strSql.AppendLine("   FROM EQUIP_CONSUME_MGT A ");

                            cmd.CommandType = CommandType.Text;
                            cmd.CommandText = strSql.ToString();

                            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                            double dTemp = 0;

                            if (dt.Rows.Count == 0)
                            {
                                dTemp = 1;
                                sConsumeCd = Convert.ToString(dTemp);
                            }
                            else
                            {
                                dTemp = Convert.ToDouble(dt.Rows[0]["MAX_VALUE"]) + 1;
                                sConsumeCd = Convert.ToString(dTemp);
                            }
                        }

                        sConItemNm = dtMerge.Rows[i]["CON_ITEM_NM"].ToString();
                        sConItemSpec = dtMerge.Rows[i]["CON_ITEM_SPEC"].ToString();
                        sConItemPurp = dtMerge.Rows[i]["CON_ITEM_PURP"].ToString();
                        sConLargeCate = dtMerge.Rows[i]["CON_LARGE_CATE"].ToString();
                        sExchangeCycle = dtMerge.Rows[i]["EXCHANGE_CYCLE"].ToString();
                        dOptimalStock = String.IsNullOrEmpty(dtMerge.Rows[i]["OPTIMAL_STOCK"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["OPTIMAL_STOCK"].ToString());
                        sItemUnit = dtMerge.Rows[i]["ITEM_UNIT"].ToString();
                        dStockAmt = String.IsNullOrEmpty(dtMerge.Rows[i]["STOCK_AMT"].ToString()) ? 0 : Convert.ToDouble(dtMerge.Rows[i]["STOCK_AMT"].ToString());
                        if (dOptimalStock < dStockAmt)
                        {

                        }
                        else if (dOptimalStock > dStockAmt)
                        {
                            dLackStockAmt = dOptimalStock - dStockAmt;
                        }

                        sNote = dtMerge.Rows[i]["NOTE"].ToString();
                        sUseYn = dtMerge.Rows[i]["USE_YN"].ToString();
                        
                        strSql.Clear();

                        strSql.AppendLine(" MERGE INTO EQUIP_CONSUME_MGT AS a     ");
                        strSql.AppendLine("     USING(SELECT CON_ITEM_CD = " + sConsumeCd + "     ");
                        strSql.AppendLine("                , CON_ITEM_NM = '" + sConItemNm + "'    ");
                        strSql.AppendLine("              , CON_ITEM_SPEC = '" + sConItemSpec + "'    ");
                        strSql.AppendLine("              , CON_ITEM_PURP = '" + sConItemPurp + "'    ");
                        strSql.AppendLine("              , CON_LARGE_CATE = '" + sConLargeCate + "'   ");
                        strSql.AppendLine("              , EXCHANGE_CYCLE = '" + sExchangeCycle + "'   ");
                        strSql.AppendLine("              , OPTIMAL_STOCK = " + dOptimalStock + "      ");
                        strSql.AppendLine("              , ITEM_UNIT = '" + sItemUnit + "'        ");
                        strSql.AppendLine("              , STOCK_AMT = " + dStockAmt + "         ");
                        strSql.AppendLine("              , LACK_STOCK_AMT = " + dLackStockAmt + "     ");
                        strSql.AppendLine("              , NOTE = '" + sNote + "'             ");
                        strSql.AppendLine("              , USE_YN = '" + sUseYn + "'           ");
                        strSql.AppendLine("              , ENT_DT = CONVERT(VARCHAR(20),GETDATE(),20)     ");
                        strSql.AppendLine("              , ENT_ID = '"+ sId + "'      ) AS b");
                        strSql.AppendLine("    ON a.CON_ITEM_CD = b.CON_ITEM_CD   ");
                        strSql.AppendLine("    WHEN MATCHED THEN UPDATE SET       ");
                        strSql.AppendLine("         CON_ITEM_NM = '" + sConItemNm + "'    ");
                        strSql.AppendLine("              , CON_ITEM_SPEC = '" + sConItemSpec + "'    ");
                        strSql.AppendLine("              , CON_ITEM_PURP = '" + sConItemPurp + "'    ");
                        strSql.AppendLine("              , CON_LARGE_CATE = '" + sConLargeCate + "'   ");
                        strSql.AppendLine("              , EXCHANGE_CYCLE = '" + sExchangeCycle + "'   ");
                        strSql.AppendLine("              , OPTIMAL_STOCK = " + dOptimalStock + "      ");
                        strSql.AppendLine("              , ITEM_UNIT = '" + sItemUnit + "'        ");
                        strSql.AppendLine("              , STOCK_AMT = " + dStockAmt + "         ");
                        strSql.AppendLine("              , LACK_STOCK_AMT = " + dLackStockAmt + "     ");
                        strSql.AppendLine("              , NOTE = '" + sNote + "'             ");
                        strSql.AppendLine("              , USE_YN = '" + sUseYn + "'           ");
                        strSql.AppendLine("              , MFY_DT = CONVERT(VARCHAR(20),GETDATE(),20)     ");
                        strSql.AppendLine("              , MFY_ID = '" + sId + "' ");
                        strSql.AppendLine("    WHEN NOT MATCHED THEN INSERT(      ");
                        strSql.AppendLine("         CON_ITEM_NM                   ");
                        strSql.AppendLine("              , CON_ITEM_SPEC          ");
                        strSql.AppendLine("              , CON_ITEM_PURP          ");
                        strSql.AppendLine("              , CON_LARGE_CATE         ");
                        strSql.AppendLine("              , EXCHANGE_CYCLE         ");
                        strSql.AppendLine("              , OPTIMAL_STOCK          ");
                        strSql.AppendLine("              , ITEM_UNIT              ");
                        strSql.AppendLine("              , STOCK_AMT              ");
                        strSql.AppendLine("              , LACK_STOCK_AMT         ");
                        strSql.AppendLine("              , NOTE                   ");
                        strSql.AppendLine("              , USE_YN                 ");
                        strSql.AppendLine("              , ENT_DT                 ");
                        strSql.AppendLine("              , ENT_ID )                ");
                        strSql.AppendLine("    VALUES(      '" + sConItemNm + "'                         ");
                        strSql.AppendLine("              , '" + sConItemSpec + "'                    ");
                        strSql.AppendLine("              , '" + sConItemPurp + "'                    ");
                        strSql.AppendLine("              , '" + sConLargeCate + "'                    ");
                        strSql.AppendLine("              , '" + sExchangeCycle + "'                    ");
                        strSql.AppendLine("              , " + dOptimalStock + "                      ");
                        strSql.AppendLine("              , '" + sItemUnit + "'                    ");
                        strSql.AppendLine("              , " + dStockAmt + "                      ");
                        strSql.AppendLine("              , " + dLackStockAmt + "                      ");
                        strSql.AppendLine("              , '" + sNote + "'                    ");
                        strSql.AppendLine("              , '" + sUseYn + "'                    ");
                        strSql.AppendLine("              , CONVERT(VARCHAR(20),GETDATE(),20)              ");
                        strSql.AppendLine("              , '"+ sId + "'               ");
                        strSql.AppendLine("    );                                 ");


                        /*
                        strSql.AppendLine(" ");
                        strSql.AppendLine("INSERT INTO EQUIP_CONSUME_MGT ");
                        strSql.AppendLine("          ( ");
                        strSql.AppendLine("            CON_ITEM_CD ");
                        strSql.AppendLine("          , CON_ITEM_NM ");
                        strSql.AppendLine("          , CON_ITEM_SPEC ");
                        strSql.AppendLine("          , CON_ITEM_PURP ");
                        strSql.AppendLine("          , CON_LARGE_CATE");
                        strSql.AppendLine("          , EXCHANGE_CYCLE ");
                        strSql.AppendLine("          , OPTIMAL_STOCK ");
                        strSql.AppendLine(" 	     , ITEM_UNIT ");
                        strSql.AppendLine(" 	     , STOCK_AMT ");
                        strSql.AppendLine(" 	     , LACK_STOCK_AMT ");
                        strSql.AppendLine(" 	     , NOTE ");
                        strSql.AppendLine(" 	     , USE_YN ");
                        strSql.AppendLine(" 	     , ENT_DT ");
                        strSql.AppendLine(" 	     , ENT_ID ");
                        strSql.AppendLine(" 	     , MFY_DT ");
                        strSql.AppendLine(" 	     , MFY_ID ");
                        strSql.AppendLine(" 	     ) ");
                        strSql.AppendLine("     VALUES ");
                        strSql.AppendLine("          ( ");
                        strSql.AppendLine("            " + sConsumeCd + " ");
                        strSql.AppendLine(" 		 , '" + sConItemNm + "'   ");
                        strSql.AppendLine(" 		 , '" + sConItemSpec + "'   ");
                        strSql.AppendLine(" 		 , '" + sConItemPurp + "'   ");
                        strSql.AppendLine(" 	     , '" + sConLargeCate + "'     ");
                        strSql.AppendLine(" 		 , '" + sExchangeCycle + "'    ");
                        strSql.AppendLine(" 		 , " + dOptimalStock + " ");
                        strSql.AppendLine(" 		 , '" + sItemUnit + "'   ");
                        strSql.AppendLine(" 		 , " + dStockAmt + " ");
                        strSql.AppendLine(" 		 , " + dLackStockAmt + " ");
                        strSql.AppendLine(" 		 , '" + sNote + "' ");
                        strSql.AppendLine(" 		 , '" + sUseYn + "' ");
                        strSql.AppendLine(" 		 , NOW() ");
                        strSql.AppendLine(" 		 , '" + sEntId + "' ");
                        strSql.AppendLine(" 		 , NOW() ");
                        strSql.AppendLine(" 		 , '" + sMfyId + "' ");
                        strSql.AppendLine("          ) ");
                        strSql.AppendLine("         ON DUPLICATE KEY UPDATE ");
                        strSql.AppendLine("            CON_ITEM_NM = '" + sConItemNm + "' ");
                        strSql.AppendLine("          , CON_ITEM_SPEC = '" + sConItemSpec + "' ");
                        strSql.AppendLine("          , CON_ITEM_PURP = '" + sConItemPurp + "' ");
                        strSql.AppendLine("          , CON_LARGE_CATE = '" + sConLargeCate + "' ");
                        strSql.AppendLine("          , EXCHANGE_CYCLE = '" + sExchangeCycle + "' ");
                        strSql.AppendLine("          , OPTIMAL_STOCK = " + dOptimalStock + " ");
                        strSql.AppendLine(" 	     , ITEM_UNIT = '" + sItemUnit + "' ");
                        strSql.AppendLine(" 	     , STOCK_AMT = " + dStockAmt + " ");
                        strSql.AppendLine(" 	     , LACK_STOCK_AMT = " + dLackStockAmt + " ");
                        strSql.AppendLine(" 	     , NOTE = '" + sNote + "' ");
                        strSql.AppendLine(" 	     , USE_YN = '" + sUseYn + "' ");
                        strSql.AppendLine(" 	     , MFY_DT = NOW() ");
                        strSql.AppendLine(" 	     , MFY_ID = '" + sMfyId + "' ");
                        */


                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();

                        string sLogRmk = "Table:EQUIP_CONSUME_MGT -> CON_ITEM_CD:" + sConsumeCd + "CON_ITEM_NM" + sConItemNm;
                        ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, (i+1).ToString(), "S", this.Name, sLogRmk, cmd);
                    }

                    DBConn.dbTran.Commit();
                    DBConn.dbTran = null;
                    XtraMessageBox.Show("저장을 완료했습니다.");

                    //DataTable dtTemp = (DataTable)GridRetr.DataSource;
                    //dtTemp.Rows.Clear();
                    //GridRetr.DataSource = dtTemp;

                    BtnRetr_Click(null, null);
                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColConItemNm, dtMerge.Rows[0]["CON_ITEM_NM"]?.ToString());
                }
                else
                {
                    XtraMessageBox.Show("변경내역이 없습니다.");
                }

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

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("DELETE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sConCd = GridViewRetr.GetFocusedRowCellValue("CON_ITEM_CD")?.ToString();
            string sConNm = GridViewRetr.GetFocusedRowCellValue("CON_ITEM_NM")?.ToString();
            if (string.IsNullOrEmpty(sConCd))
            {
                XtraMessageBox.Show("소모품코드가 존재하지 않습니다. \r\n 삭제하려는 데이터를 정확히 클릭하세요.");
                return;
            }

            if (XtraMessageBox.Show("소모품코드 : " + sConCd + "\r\n소모품명 : " + sConNm + " \r\n선택된 항목을 삭제하시겠습니까? \r\n 삭제한 데이터는 복구할 수 없습니다."
                , "소모품 항목 삭제여부", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }

            if(string.IsNullOrEmpty(sConCd) || string.IsNullOrEmpty(sConNm))
            {
                XtraMessageBox.Show("삭제 필수사항이 존재하지 않습니다.");
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
                strSql.AppendLine(" DELETE FROM EQUIP_CONSUME_MGT");
                strSql.AppendLine("       WHERE CON_ITEM_CD = " + sConCd + " ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                string sLogRmk = "Table:EQUIP_CONSUME_MGT -> CON_ITEM_CD:" + sConCd;
                ClsFunc.LogInsert(DateTime.Now.ToString(), FmMainToolBar2.UserID, "1", "D", this.Name, sLogRmk, cmd);

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("삭제 완료했습니다.");

                int idx = GridViewRetr.FocusedRowHandle;

                BtnRetr_Click(null, null);

                GridViewRetr.FocusedRowHandle = idx - 1;
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

        private void GridViewRetr_KeyDown(object sender, KeyEventArgs e)
        {
            //if (e.KeyCode == Keys.Enter)
            //{
            //    GridViewRetr.OptionsBehavior.EditingMode = GridEditingMode.EditFormInplace;
            //    GridViewRetr.ShowEditForm();
            //}
        }

        private void GridViewRetr_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            //GridViewRetr.OptionsBehavior.EditingMode = GridEditingMode.Default;
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
                string sFileNM = "소모품리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;

                    GridRetr.ExportToXls(FileName + ".xls");
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

        private void GridViewRetr_RowCellStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs e)
        {
            GridView view = sender as GridView;

            if (e.Column.FieldName == "LACK_STOCK_AMT")
            {
                string sOptmAmt = view.GetRowCellDisplayText(e.RowHandle, view.Columns["OPTIMAL_STOCK"]).ToString();
                string sLackAmt = view.GetRowCellDisplayText(e.RowHandle, view.Columns["LACK_STOCK_AMT"]).ToString();
                string sStockAmt = view.GetRowCellDisplayText(e.RowHandle, view.Columns["STOCK_AMT"]).ToString();

                double dOptmAMt = 0;
                double dStockAmt = 0;
                double dLackAmt = 0;

                double.TryParse(sOptmAmt, out dOptmAMt);
                double.TryParse(sStockAmt, out dStockAmt);

                dLackAmt = dStockAmt - dOptmAMt;

                if (dLackAmt < 0) //dLackAmt < dOptmAMt || dStockAmt < 0
                {
                    e.Appearance.BackColor = Color.IndianRed;
                }
            }
        }
        
        private void LkupMenuItemNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void GridViewRetr_ShowingPopupEditForm(object sender, ShowingPopupEditFormEventArgs e)
        {
            e.EditForm.ImeMode = ImeMode.Hangul;
            e.EditForm.StartPosition = FormStartPosition.CenterParent;

            foreach (var button in e.EditForm.Controls.OfType<EditFormContainer>()
                        .SelectMany(control => Enumerable.Cast<Control>(control.Controls)).OfType<PanelControl>()
                        .SelectMany(nestedControl => Enumerable.Cast<Control>(nestedControl.Controls)).OfType<SimpleButton>())
            {
                switch (button.Text)
                {
                    case "Update":
                        //button.Visible = false;
                        button.Text = "수정";
                        //button.Click += EditFormUpdateButton_Click;
                        break;

                    case "Cancel":
                        button.Text = "취소";
                        break;
                }
            }
        }

        private void BtnInOutRgt_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sCON_ITEM_CD = GridViewRetr.GetFocusedRowCellValue("CON_ITEM_CD")?.ToString();
            if (!string.IsNullOrEmpty(sCON_ITEM_CD))
            {
                EquipConsumeInOutMgt frm = new EquipConsumeInOutMgt();
                frm.sConsumeCd = sCON_ITEM_CD;

                if (frm.ShowDialog() == DialogResult.OK)
                {
                    BtnRetr_Click(null, null);
                    GridViewRetr.FocusedRowHandle = GridViewRetr.LocateByDisplayText(0, GridColConItemCd, sCON_ITEM_CD);
                }
            }
            else
            {
                XtraMessageBox.Show("해당 소모품에 대한 정보가 없습니다. \r\n 저장 버튼을 누른 후 진행해주세요.");
                return;
            }
        }

        private void EquipConsumeCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnNew_Click(null, null);
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
            else if (e.KeyCode == Keys.F8)
            {
                BtnExcel_Click(null, null);
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, RowStyleEventArgs e)
        {
            ComnEventFunc.SettingGridViewRowPatternToStripe(sender, e);
        }

        private void EquipConsumeCd_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void EquipConsumeCd_TextChanged(object sender, EventArgs e)
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
    }
}