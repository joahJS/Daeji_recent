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
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class EquipConsumeInOutMgt : DevExpress.XtraEditors.XtraForm
    {
        public EquipConsumeInOutMgt()
        {
            InitializeComponent();
        }

        public string sConsumeCd { get; set; }

        private void EquipConsumeInOutMgt_Load(object sender, EventArgs e)
        {
            DataTable dtInOutGb = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoConsumeGLkupInOutGb, dtInOutGb, GridHistory, GridColConInOutGb, "CD", "NM", "");

            DataTable dtCv = GetLookUpData("4", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoLkupCv, dtCv, GridHistory, GridColConSuplDealer, "CD", "NM", "");

            GetEquipConsumeData();
        }

        private void GetEquipConsumeData()
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT CON_ITEM_CD    ");
            strSql.AppendLine("      , 'MOD' AS EDIT");
            strSql.AppendLine("      , CON_ITEM_SEQ   ");
            strSql.AppendLine("      , ITEM_INOUT_GB  ");
            strSql.AppendLine("      , CONVERT(DATE,OCCUR_DT) AS OCCUR_DT       ");
            strSql.AppendLine("      , CON_ITEM_AMT   ");
            strSql.AppendLine("      , CON_ITEM_UNPR  ");
            strSql.AppendLine("      , CON_ITEM_PRICE ");
            strSql.AppendLine("      , CON_SUPL_DEALER");
            strSql.AppendLine("      , CON_STORAGE    ");
            strSql.AppendLine("      , NOTE           ");
            strSql.AppendLine("      , '' AS IDT_NO          ");
            strSql.AppendLine("      , ENT_DT         ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(ENT_ID AS NUMERIC) IS NULL THEN ENT_ID ELSE DBO.FN_USRNM(ENT_ID) END AS ENT_ID ");
            strSql.AppendLine("      , MFY_DT                                                                                             ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(MFY_ID  AS NUMERIC) IS NULL THEN MFY_ID ELSE DBO.FN_USRNM(MFY_ID) END AS MFY_ID");
            strSql.AppendLine("   FROM EQUIP_CONSUME_HISTORY                                                                              ");
            strSql.AppendLine("  WHERE CON_ITEM_CD = " + sConsumeCd);
            strSql.AppendLine("  ORDER BY CON_ITEM_SEQ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridHistory.DataSource = dt;
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
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE CD_GB = 'ITEM_INOUT_GB'");
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
                strSql.AppendLine(" SELECT CONVERT(VARCHAR,DEALER_CD) AS CD ");
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

        private void BtnHisSave_Click(object sender, EventArgs e)
        {
            DataTable dtChk = (DataTable)GridHistory.DataSource;
            if(dtChk == null)
            {
                XtraMessageBox.Show("저장할 데이터가 없습니다.");
                return;
            }

            DataSet saveDT = ComGrid.GET_DATASET_FOR_SAVE(dtChk);
            DataTable addDt = saveDT.Tables[0];

            if (addDt.Rows.Count == 0)
            {
                XtraMessageBox.Show("저장할 데이터가 없습니다.");
                return;
            }

            for (int i = 0; i < addDt.Rows.Count; i++)
            {
                if (String.IsNullOrEmpty(addDt.Rows[i]["ITEM_INOUT_GB"].ToString()))
                {
                    XtraMessageBox.Show("입출고 구분을 입력하세요.");
                    //GridViewHistory.FocusedRowHandle = i;
                    return;
                }
                if (String.IsNullOrEmpty(addDt.Rows[i]["OCCUR_DT"].ToString()))
                {
                    XtraMessageBox.Show("발생일자를 입력하세요.");
                    //GridViewHistory.FocusedRowHandle = i;
                    return;
                }
                if (String.IsNullOrEmpty(addDt.Rows[i]["CON_ITEM_AMT"].ToString()))
                {
                    XtraMessageBox.Show("소모품 수량을 입력하세요.");
                    //GridViewHistory.FocusedRowHandle = i;
                    return;
                }
                if (addDt.Rows[i]["ITEM_INOUT_GB"].ToString() == "I")
                {
                    if (String.IsNullOrEmpty(addDt.Rows[i]["CON_ITEM_UNPR"].ToString()))
                    {
                        XtraMessageBox.Show("소모품 단가를 입력하세요.");
                        //GridViewHistory.FocusedRowHandle = i;
                        return;
                    }
                    if (String.IsNullOrEmpty(addDt.Rows[i]["CON_ITEM_PRICE"].ToString()))
                    {
                        XtraMessageBox.Show("소모품 가격을 입력하세요.");
                        //GridViewHistory.FocusedRowHandle = i;
                        return;
                    }
                    if (String.IsNullOrEmpty(addDt.Rows[i]["CON_SUPL_DEALER"].ToString()))
                    {
                        XtraMessageBox.Show("공급업체를 입력하세요.");
                        //GridViewHistory.FocusedRowHandle = i;
                        return;
                    }
                }
                if (String.IsNullOrEmpty(addDt.Rows[i]["CON_STORAGE"].ToString()))
                {
                    XtraMessageBox.Show("보관장소를 입력하세요.");
                    //GridViewHistory.FocusedRowHandle = i;
                    return;
                }

            }

            try
            {
                Cursor = Cursors.WaitCursor;

                double dConItemCd = 0;
                double dConItemSeq = 0;
                string sItemIOGb = string.Empty;
                string sOccurDt = string.Empty;
                double dConItemAmt = 0;
                double dConItemUnpr = 0;
                double dConItemPrc = 0;
                string sConSuplDealer = string.Empty;
                string sConStorage = string.Empty;
                string sNote = string.Empty;
                string sId = FmMainToolBar2.drUser["USRCD"]?.ToString();

                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                DataTable dt = addDt;
                StringBuilder strSql = new StringBuilder();

                double dSumAmt = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    double dResult = 0;
                    if (String.IsNullOrEmpty(dt.Rows[i]["CON_ITEM_SEQ"].ToString()))
                    {
                        strSql.Clear();
                        strSql.AppendLine(" ");
                        strSql.AppendLine(" SELECT MAX(A.CON_ITEM_SEQ) AS MAX_SEQ ");
                        strSql.AppendLine("   FROM EQUIP_CONSUME_HISTORY A");
                        strSql.AppendLine("  WHERE CON_ITEM_CD = " + dt.Rows[0]["CON_ITEM_CD"].ToString() + " ");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        dResult = String.IsNullOrEmpty(cmd.ExecuteScalar().ToString()) ? 0 : Convert.ToDouble(cmd.ExecuteScalar().ToString());

                        if (dResult == 0)
                        {
                            dResult = 1;
                        }
                        else
                        {
                            dResult++;
                        }
                    }

                    dConItemCd = Convert.ToDouble(dt.Rows[i]["CON_ITEM_CD"].ToString());
                    dConItemSeq = dResult;
                    sItemIOGb = Convert.ToString(dt.Rows[i]["ITEM_INOUT_GB"]);
                    sOccurDt = dt.Rows[i]["OCCUR_DT"].ToString().Replace("-", "").Substring(0, 8);
                    dConItemAmt = String.IsNullOrEmpty(dt.Rows[i]["CON_ITEM_AMT"].ToString()) ? 0 : Convert.ToDouble(dt.Rows[i]["CON_ITEM_AMT"].ToString());
                    dConItemUnpr = String.IsNullOrEmpty(dt.Rows[i]["CON_ITEM_UNPR"].ToString()) ? 0 : Convert.ToDouble(dt.Rows[i]["CON_ITEM_UNPR"].ToString());
                    dConItemPrc = String.IsNullOrEmpty(dt.Rows[i]["CON_ITEM_PRICE"].ToString()) ? 0 : Convert.ToDouble(dt.Rows[i]["CON_ITEM_PRICE"].ToString());
                    sConSuplDealer = Convert.ToString(dt.Rows[i]["CON_SUPL_DEALER"]);
                    sConStorage = Convert.ToString(dt.Rows[i]["CON_STORAGE"]);
                    sNote = Convert.ToString(dt.Rows[i]["NOTE"]);
                    
                    strSql.Clear();
                    strSql.AppendLine(" IF EXISTS(SELECT* FROM EQUIP_CONSUME_HISTORY WHERE CON_ITEM_CD = "+ dConItemCd + " AND CON_ITEM_SEQ = "+ dConItemSeq + ")");
                    strSql.AppendLine("    BEGIN                                                                                ");
                    strSql.AppendLine("          UPDATE EQUIP_CONSUME_HISTORY                                                   ");
                    strSql.AppendLine("             SET ITEM_INOUT_GB =    '" + sItemIOGb + "' "); ;
                    strSql.AppendLine("               , OCCUR_DT =         '" + sOccurDt + "' ");                                            
                    strSql.AppendLine("               , CON_ITEM_AMT =     " + dConItemAmt + " ");                                           
                    strSql.AppendLine("               , CON_ITEM_UNPR =    " + dConItemUnpr + " ");                                          
                    strSql.AppendLine("               , CON_ITEM_PRICE =   " + dConItemPrc + " ");                                           
                    strSql.AppendLine("               , CON_SUPL_DEALER =  '" + sConSuplDealer + "' ");                                      
                    strSql.AppendLine("               , CON_STORAGE =      '" + sConStorage + "' ");                                         
                    strSql.AppendLine("               , NOTE =             '" + sNote + "' ");                                               
                    strSql.AppendLine("               , MFY_DT = CONVERT(VARCHAR(20), GETDATE(), 20)                            ");
                    strSql.AppendLine("               , MFY_ID =  '"+ sId + "'                                                  ");
                    strSql.AppendLine("           WHERE CON_ITEM_CD = " + dConItemCd + " AND CON_ITEM_SEQ = " + dConItemSeq      );
                    strSql.AppendLine("      END                                                                                ");
                    strSql.AppendLine(" ELSE                                                                                    ");
                    strSql.AppendLine("    BEGIN                                                                                ");
                    strSql.AppendLine("          INSERT INTO EQUIP_CONSUME_HISTORY ");
                    strSql.AppendLine("                    ( CON_ITEM_CD ");
                    strSql.AppendLine("                    , CON_ITEM_SEQ ");
                    strSql.AppendLine("                    , ITEM_INOUT_GB ");
                    strSql.AppendLine("                    , OCCUR_DT ");
                    strSql.AppendLine("                    , CON_ITEM_AMT ");
                    strSql.AppendLine("                    , CON_ITEM_UNPR ");
                    strSql.AppendLine("                    , CON_ITEM_PRICE ");
                    strSql.AppendLine("                    , CON_SUPL_DEALER ");
                    strSql.AppendLine("                    , CON_STORAGE ");
                    strSql.AppendLine("                    , NOTE ");
                    strSql.AppendLine("                    , ENT_DT ");
                    strSql.AppendLine("                    , ENT_ID ");
                    strSql.AppendLine("                    ) ");
                    strSql.AppendLine("               VALUES ");
                    strSql.AppendLine("                    ( " + dConItemCd + " ");
                    strSql.AppendLine("                    , " + dConItemSeq + " ");
                    strSql.AppendLine("                    , '" + sItemIOGb + "' ");
                    strSql.AppendLine("                    , '" + sOccurDt + "' ");
                    strSql.AppendLine("                    , " + dConItemAmt + " ");
                    strSql.AppendLine("                    , " + dConItemUnpr + " ");
                    strSql.AppendLine("                    , " + dConItemPrc + " ");
                    strSql.AppendLine("                    , '" + sConSuplDealer + "' ");
                    strSql.AppendLine("                    , '" + sConStorage + "' ");
                    strSql.AppendLine("                    , '" + sNote + "' ");
                    strSql.AppendLine("                    , CONVERT(VARCHAR(20), GETDATE(), 20) ");
                    strSql.AppendLine("                    , '" + sId + "' ");
                    strSql.AppendLine("           ) ");
                    strSql.AppendLine("      END                                                                                ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();

                    if (dt.Rows[i]["ITEM_INOUT_GB"].ToString().Equals("I"))
                    {
                        dSumAmt += String.IsNullOrEmpty(dt.Rows[i]["CON_ITEM_AMT"].ToString()) ? 0 : Convert.ToDouble(dt.Rows[i]["CON_ITEM_AMT"].ToString());
                    }
                    else if (dt.Rows[i]["ITEM_INOUT_GB"].ToString().Equals("O"))
                    {
                        dSumAmt -= String.IsNullOrEmpty(dt.Rows[i]["CON_ITEM_AMT"].ToString()) ? 0 : Convert.ToDouble(dt.Rows[i]["CON_ITEM_AMT"].ToString());
                    }
                }

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT A.STOCK_AMT ");
                strSql.AppendLine("      , A.OPTIMAL_STOCK ");
                strSql.AppendLine("   FROM EQUIP_CONSUME_MGT A");
                strSql.AppendLine("  WHERE CON_ITEM_CD = " + dt.Rows[0]["CON_ITEM_CD"] + " ");

                DataTable dtResult = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                double dStockAmt = String.IsNullOrEmpty(dtResult.Rows[0]["STOCK_AMT"].ToString()) ? 0 : Convert.ToDouble(dtResult.Rows[0]["STOCK_AMT"].ToString());
                double dOptmStockAmt = String.IsNullOrEmpty(dtResult.Rows[0]["OPTIMAL_STOCK"].ToString()) ? 0 : Convert.ToDouble(dtResult.Rows[0]["OPTIMAL_STOCK"].ToString());
                double dResultAmt = (dStockAmt + dSumAmt);
                double dLackAmt = 0;

                if (dResultAmt > dOptmStockAmt)
                {
                    dLackAmt = 0;
                }
                else if (dResultAmt < dOptmStockAmt)
                {
                    dLackAmt = dResultAmt - dOptmStockAmt;
                    dLackAmt = Math.Abs(dLackAmt);
                }

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" UPDATE EQUIP_CONSUME_MGT ");
                strSql.AppendLine("    SET STOCK_AMT = " + dResultAmt + " ");
                strSql.AppendLine("   , LACK_STOCK_AMT = " + dLackAmt + " ");
                strSql.AppendLine("  WHERE CON_ITEM_CD = " + dt.Rows[0]["CON_ITEM_CD"] + " ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                XtraMessageBox.Show("저장을 완료했습니다.");

                DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void GridViewHistory_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            string sEdit = GridViewHistory.GetFocusedRowCellValue("EDIT")?.ToString();

            if (string.IsNullOrEmpty(sEdit))
            {
                GridViewHistory.OptionsBehavior.Editable = true;
            }
            else
            {
                GridViewHistory.OptionsBehavior.Editable = false;
                return;
            }

            //if (String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellDisplayText("ITEM_INOUT_GB")))
            //{
            //    RepoConsumeGLkupInOutGb.ReadOnly = false;
            //    RepoConsumeTxtUnpr.ReadOnly = false;
            //    RepoConsumeTxtPrice.ReadOnly = false;

            //    GridColConUnpr.OptionsColumn.AllowEdit = false;
            //    GridColConPrice.OptionsColumn.AllowEdit = false;
            //    GridColConSuplDealer.OptionsColumn.AllowEdit = false;
            //}
            //else 
            if (!String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellDisplayText("ITEM_INOUT_GB")))
            {
                if (GridViewHistory.GetFocusedRowCellValue("ITEM_INOUT_GB").ToString() == "O")
                {
                    RepoConsumeGLkupInOutGb.ReadOnly = false;
                    RepoConsumeTxtUnpr.ReadOnly = true;
                    RepoConsumeTxtPrice.ReadOnly = true;

                    GridColConUnpr.OptionsColumn.AllowEdit = false;
                    GridColConPrice.OptionsColumn.AllowEdit = false;
                    GridColConSuplDealer.OptionsColumn.AllowEdit = false;
                }
                else if (GridViewHistory.GetFocusedRowCellValue("ITEM_INOUT_GB").ToString() == "I")
                {
                    RepoConsumeGLkupInOutGb.ReadOnly = false;
                    RepoConsumeTxtUnpr.ReadOnly = false;
                    RepoConsumeTxtPrice.ReadOnly = false;

                    GridColConUnpr.OptionsColumn.AllowEdit = true;
                    GridColConPrice.OptionsColumn.AllowEdit = true;
                    GridColConSuplDealer.OptionsColumn.AllowEdit = true;
                }
            }
        }

        private void RepoConsumeGLkupInOutGb_Leave(object sender, EventArgs e)
        {
            string sInOutGb = (sender as GridLookUpEdit).EditValue == null ? string.Empty : (sender as GridLookUpEdit).EditValue.ToString();

            if (sInOutGb == "I")
            {
                RepoConsumeGLkupInOutGb.ReadOnly = false;
                RepoConsumeTxtUnpr.ReadOnly = false;
                RepoConsumeTxtPrice.ReadOnly = false;

                GridColConUnpr.OptionsColumn.AllowEdit = true;
                GridColConPrice.OptionsColumn.AllowEdit = true;
                GridColConSuplDealer.OptionsColumn.AllowEdit = true;
            }else if(sInOutGb == "O")
            {
                RepoConsumeGLkupInOutGb.ReadOnly = false;
                RepoConsumeTxtUnpr.ReadOnly = true;
                RepoConsumeTxtPrice.ReadOnly = true;

                GridColConUnpr.OptionsColumn.AllowEdit = false;
                GridColConPrice.OptionsColumn.AllowEdit = false;
                GridColConSuplDealer.OptionsColumn.AllowEdit = false;
            }
        }

        private void RepoConsumeTxtPrice_Leave(object sender, EventArgs e)
        {
            if ((sender as TextEdit).EditValue == null)
            {
                return;
            }

            if (String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_AMT").ToString()) &&
                String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_UNPR").ToString()))
            {
                return;
            }
            else if (!String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_AMT").ToString()))
            {
                double dItemPrc = 0;
                double.TryParse((sender as TextEdit).EditValue.ToString(), out dItemPrc);

                double dItemAmt = 0;
                double.TryParse(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_AMT").ToString(), out dItemAmt);

                GridViewHistory.SetFocusedRowCellValue("CON_ITEM_UNPR", (dItemPrc / dItemAmt));

            }
            else if (!String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_UNPR").ToString()))
            {
                double dItemPrc = 0;
                double.TryParse((sender as TextEdit).EditValue.ToString(), out dItemPrc);

                double dItemUnpr = 0;
                double.TryParse(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_UNPR").ToString(), out dItemUnpr);

                GridViewHistory.SetFocusedRowCellValue("CON_ITEM_AMT", (dItemPrc / dItemUnpr));
            }
        }

        private void RepoConsumeTxtUnpr_Leave(object sender, EventArgs e)
        {
            if ((sender as TextEdit).EditValue == null)
            {
                return;
            }

            if (String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_AMT").ToString()) &&
                String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_PRICE").ToString()))
            {
                return;
            }
            else if (!String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_AMT").ToString()))
            {
                double dItemUnpr = 0;
                double.TryParse((sender as TextEdit).EditValue.ToString(), out dItemUnpr);

                double dItemAmt = 0;
                double.TryParse(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_AMT").ToString(), out dItemAmt);

                GridViewHistory.SetFocusedRowCellValue("CON_ITEM_PRICE", (dItemUnpr * dItemAmt));
            }
            else if (!String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_PRICE").ToString()))
            {
                double dItemUnpr = 0;
                double.TryParse((sender as TextEdit).EditValue.ToString(), out dItemUnpr);

                double dItemPrice = 0;
                double.TryParse(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_PRICE").ToString(), out dItemPrice);

                GridViewHistory.SetFocusedRowCellValue("CON_ITEM_AMT", (dItemPrice / dItemUnpr));
            }
        }

        private void RepoConsumeTxtAmt_Leave(object sender, EventArgs e)
        {
            if ((sender as TextEdit).EditValue == null)
            {
                return;
            }
            
            if (String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_UNPR").ToString()) &&
                String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_PRICE").ToString()))
            {

            }
            else if (!String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_UNPR").ToString()))
            {
                double dItemAmt = 0;
                double.TryParse((sender as TextEdit).EditValue.ToString(), out dItemAmt);

                double dItemUnpr = 0;
                double.TryParse(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_UNPR").ToString(), out dItemUnpr);

                GridViewHistory.SetFocusedRowCellValue("CON_ITEM_PRICE", (dItemAmt * dItemUnpr));
            }
            else if (!String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_PRICE").ToString()))
            {
                double dItemAmt = 0;
                double.TryParse((sender as TextEdit).EditValue.ToString(), out dItemAmt);

                double dItemPrice = 0;
                double.TryParse(GridViewHistory.GetFocusedRowCellValue("CON_ITEM_PRICE").ToString(), out dItemPrice);

                GridViewHistory.SetFocusedRowCellValue("CON_ITEM_UNPR", (dItemPrice / dItemAmt));
            }

            if (!String.IsNullOrEmpty(GridViewHistory.GetFocusedRowCellValue("ITEM_INOUT_GB").ToString()))
            {
                if (GridViewHistory.GetFocusedRowCellValue("ITEM_INOUT_GB").ToString().Equals("O"))
                {
                    GridViewHistory.FocusedColumn = GridViewHistory.Columns[7];
                }
            }
        }

        private void RepoConsumeBtnEditSuplDealer_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            PopUpDealerCd frm = new PopUpDealerCd();

            DataRow drDealerInfo;

            if (frm.ShowDialog() == DialogResult.OK)
            {
                drDealerInfo = frm.drResult;
                (sender as ButtonEdit).Text = drDealerInfo["DEALER_NM"].ToString();
                GridViewHistory.SetFocusedRowCellValue("IDT_NO", drDealerInfo["IDT_NO"].ToString());
            }
        }

        private void BtnHisNew_Click(object sender, EventArgs e)
        {
            if(GridViewHistory.RowCount > 0)
            {
                string sInoutGb = GridViewHistory.GetRowCellValue(GridViewHistory.RowCount - 1, GridColConInOutGb)?.ToString();

                if (string.IsNullOrEmpty(sInoutGb))
                {
                    XtraMessageBox.Show("입출고 구분을 입력해주세요.");
                    GridViewHistory.FocusedRowHandle = GridViewHistory.RowCount - 1;
                    GridViewHistory.FocusedColumn = GridColConInOutGb;
                    GridViewHistory.Focus();

                    return;
                }
            }

            GridViewHistory.AddNewRow();
            GridViewHistory.SetFocusedRowCellValue("CON_ITEM_CD", sConsumeCd);
            GridViewHistory.SetFocusedRowCellValue("OCCUR_DT", DateTime.Now.ToString().Substring(0, 10));
        }

        private void EquipConsumeInOutMgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnHisNew_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnHisSave_Click(null, null);
            }
            else if (e.KeyCode == Keys.F4)
            {

            }
            else if (e.KeyCode == Keys.F5)
            {
                //BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                Cursor = Cursors.WaitCursor;
                //BtnReport_Click(null, null);
                Cursor = Cursors.Default;
            }
        }

        private void GridHistory_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewHistory_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void GridViewHistory_CellValueChanged(object sender, DevExpress.XtraGrid.Views.Base.CellValueChangedEventArgs e)
        {
            GridViewHistory.UpdateCurrentRow();
        }

        private void RepoConsumeDateEditOccurDt_EditValueChanged(object sender, EventArgs e)
        {
            DateEdit dateEdit = (DateEdit)sender;

            if(DateTime.TryParse(dateEdit.EditValue?.ToString(), out DateTime result))
            {
                GridViewHistory.SetRowCellValue(GridViewHistory.FocusedRowHandle, GridViewHistory.FocusedColumn, result.ToString("yyyy-MM-dd"));
            }
        }
    }
}