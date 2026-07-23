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
using DevExpress.XtraGrid;
using DevExpress.XtraEditors.Repository;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class MesGradeCdDev : DevExpress.XtraEditors.XtraForm
    {
        public MesGradeCdDev()
        {
            InitializeComponent();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public DataRow rowUserInfo { get; set; }
        private void MesGradeCdDev_Load(object sender, EventArgs e)
        {   
            SetLookUpEdit(LkupEditGb, "1", "Y", "Y", "");

            DataTable dt = GetLookUpData("1", "N", "");

            RepositoryItemLookUpEdit columnEditor = new RepositoryItemLookUpEdit();
            columnEditor.DataSource = dt;
            columnEditor.ValueMember = "CD";
            columnEditor.DisplayMember = "NM";

            GridRetr.RepositoryItems.Add(columnEditor);

            GridColGbCd.ColumnEdit = columnEditor;
            columnEditor.NullText = "";

            DataTable dt1 = GetLookUpData("2", "N", "****");

            RepositoryItemLookUpEdit columnEditor1 = new RepositoryItemLookUpEdit();
            columnEditor1.DataSource = dt1;
            columnEditor1.ValueMember = "CD";
            columnEditor1.DisplayMember = "NM";

            GridRetr.RepositoryItems.Add(columnEditor1);
            GridColDetCd.ColumnEdit = columnEditor1;
            columnEditor1.NullText = "";

            GridColItemCd.OptionsColumn.ReadOnly = true;

            
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, "MesGradeCdDev");
        }


        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam)
        {
            //string strSql = string.Empty;

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("Y"))
            {                
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }

            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = '0002'");
            }
            else
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER (ORDER BY A.COM_SUB_CD1 , A.SORT_SEQ) AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = '0003'");
                strSql.AppendLine("    AND (('****' = '" + sParam + "') OR (('****' <> '" + sParam + "') AND A.COM_SUB_CD1 = '" + sParam + "'))");
            }
            strSql.AppendLine(") ");
            strSql.AppendLine("    SELECT CD, NM FROM ITEM_INFO");
            strSql.AppendLine("     ORDER BY SEQ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private void SetLookUpEdit(DevExpress.XtraEditors.LookUpEdit lkup, string sGb, string sNullYn, string sSetIdx, string sParam)
        {
            DataTable dt = GetLookUpData(sGb, sNullYn, sParam);

            //string strSql = string.Empty;

            //if (sNullYn.Equals("Y"))
            //{
            //    strSql.AppendLine("WITH ITEM_INFO AS (");
            //    strSql.AppendLine("SELECT '****' AS CD");
            //    strSql.AppendLine("     , '전체' AS NM");
            //    strSql.AppendLine("     , 0 AS SEQ");
            //    strSql.AppendLine(" UNION ALL");
            //}

            //if (sGb.Equals("1"))
            //{
            //    strSql.AppendLine("SELECT A.COM_CD AS CD");
            //    strSql.AppendLine("     , A.COM_NM AS NM");
            //    strSql.AppendLine("     , A.SORT_SEQ AS SEQ");
            //    strSql.AppendLine("  FROM COM_BASE_CD A");
            //    strSql.AppendLine(" WHERE A.CD_GB = '0002'";
            //}
            //else
            //{
            //    strSql.AppendLine("SELECT A.COM_CD AS CD");
            //    strSql.AppendLine("     , A.COM_NM AS NM");
            //    strSql.AppendLine("     , ROW_NUMBER() OVER (ORDER BY A.COM_SUB_CD1 , A.SORT_SEQ) AS SEQ");
            //    strSql.AppendLine("  FROM COM_BASE_CD A");
            //    strSql.AppendLine(" WHERE A.CD_GB = '0003'";
            //    strSql.AppendLine("   AND (('****' = '" + sParam + "') OR (('****' <> '" + sParam + "') AND A.COM_SUB_CD1 = '" + sParam + "'))";                
            //}

            //strSql.AppendLine("   )");
            //strSql.AppendLine("   SELECT CD, NM FROM ITEM_INFO";
            //strSql.AppendLine(" ORDER BY SEQ");

            //DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql);

            lkup.Properties.DataSource = dt;
            lkup.Properties.DisplayMember = "NM";
            lkup.Properties.ValueMember = "CD";

            if (sSetIdx.Equals("Y")) lkup.ItemIndex = 0;
        }

        private void LkupEditGb_EditValueChanged(object sender, EventArgs e)
        {
            SetLookUpEdit(LkupEditDetGb, "2", "Y", "Y", LkupEditGb.EditValue.ToString());

            LkupEditDetGb_EditValueChanged(null, null);
        }

        private void LkupEditDetGb_EditValueChanged(object sender, EventArgs e)
        {
            GetGridRetr();
        }

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }
            GetGridRetr();
        }

        private void GetGridRetr()
        {
            GridRetr.DataSource = null;
            
            string sSubCd1 = LkupEditGb.EditValue.ToString();
            string sSubCd2 = LkupEditDetGb.EditValue.ToString();

            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.COM_SUB_CD1 AS GB_CD ");
            strSql.AppendLine("      , CONCAT(A.COM_SUB_CD1, A.COM_SUB_CD2) AS DET_CD");
            strSql.AppendLine("      , A.COM_CD AS ITEM_CD ");
            strSql.AppendLine("      , A.COM_NM AS ITEM_NM ");
            strSql.AppendLine("      , A.REMARK AS RMK       ");
            strSql.AppendLine("      , A.USE_YN       ");
            strSql.AppendLine("      , A.SORT_SEQ       ");
            strSql.AppendLine("   FROM COM_BASE_CD A ");
            strSql.AppendLine("  WHERE A.CD_GB = '0001' ");
            strSql.AppendLine("    AND (('****' = '" + sSubCd1 + "') OR(('****' <> '" + sSubCd1 + "') AND A.COM_SUB_CD1 = '" + sSubCd1 + "'))");
            strSql.AppendLine("    AND (('****' = '" + sSubCd2 + "') OR(('****' <> '" + sSubCd2 + "') AND A.COM_SUB_CD2 = RIGHT('" + sSubCd2 + "', 1)))");
            strSql.AppendLine("  ORDER BY A.COM_SUB_CD1, A.COM_SUB_CD2, A.SORT_SEQ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridRetr.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            DataTable dt = (DataTable)GridRetr.DataSource;
            

            DataSet dsSave =  ComGrid.GET_DATASET_FOR_SAVE(dt);
            DataTable dtAdd = dsSave.Tables[0];
            DataTable dtMod = dsSave.Tables[1];

            string sCdGbNm = "항목등급코드(MES)";
            string sId = "MIG";

            string sComCd = string.Empty;
            string sComSubCd1 = string.Empty;
            string sComSubCd2 = string.Empty;
            string sComSubCd3 = string.Empty;
            string sComNm = string.Empty;
            string sComSubNm1 = string.Empty;
            string sComSubNm2 = string.Empty;
            string sComSubNm3 = string.Empty;
            string sRemark = string.Empty;
            string sUseYn = string.Empty;
            string sSortSeq = string.Empty;
            
            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            StringBuilder strSql = new StringBuilder();

            try
            {
                if (dtMod.Rows.Count > 0)  // modify
                {
                    for (int j = 0; j < dtMod.Rows.Count; j++)
                    {
                        sComCd = Convert.ToString(dtMod.Rows[j]["ITEM_CD"]);
                        sComNm = Convert.ToString(dtMod.Rows[j]["ITEM_NM"]);
                        sComSubNm3 = Convert.ToString(dtMod.Rows[j]["ITEM_NM"]);
                        sRemark = Convert.ToString(dtMod.Rows[j]["RMK"]);
                        sUseYn = Convert.ToString(dtMod.Rows[j]["USE_YN"]);
                        sSortSeq = Convert.ToString(dtMod.Rows[j]["SORT_SEQ"]);

                        strSql.AppendLine(" ");
                        strSql.AppendLine(" UPDATE COM_BASE_CD");
                        strSql.AppendLine("    SET COM_NM = '" + sComNm + "'");
                        strSql.AppendLine("      , COM_SUB_NM3 = '" + sComSubNm3 + "'");
                        strSql.AppendLine("      , REMARK = '" + sRemark + "'");
                        strSql.AppendLine("      , USE_YN = '" + sUseYn + "'");
                        strSql.AppendLine("      , SORT_SEQ = '" + sSortSeq + "'");
                        strSql.AppendLine("      , MFY_DT = NOW()");
                        strSql.AppendLine("      , MFY_ID = '" + sId + "'");
                        strSql.AppendLine("  WHERE CD_GB = '0001'");
                        strSql.AppendLine("    AND COM_CD = '" + sComCd + "'");

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }
                }

                //string sComCd = GridViewRetr.GetFocusedRowCellValue("ITEM_CD").ToString();
                //string sComSubCd1 = GridViewRetr.GetFocusedRowCellValue("GB_CD").ToString();
                //string sComSubCd2 = GridViewRetr.GetFocusedRowCellValue("DET_CD").ToString().Substring(1,1);
                //string sComSubCd3 = sComCd.Substring(2, 2);
                //string sComNm = GridViewRetr.GetFocusedRowCellValue("ITEM_NM").ToString();
                //string sComSubNm1 = GridViewRetr.GetFocusedRowCellDisplayText("GB_CD").ToString();
                //string sComSubNm2 = GridViewRetr.GetFocusedRowCellDisplayText("DET_CD").ToString();
                //string sComSubNm3 = GridViewRetr.GetFocusedRowCellDisplayText("ITEM_CD").ToString();
                //string sRemark = GridViewRetr.GetFocusedRowCellValue("RMK").ToString();
                //string sUseYn = GridViewRetr.GetFocusedRowCellValue("USE_YN").ToString();
                //string sSortSeq = GridViewRetr.GetFocusedRowCellValue("SORT_SEQ").ToString();

                if (dtAdd.Rows.Count > 0)   // insert
                {

                    for (int j = 0; j < dtAdd.Rows.Count; j++)
                    {
                        sComNm = Convert.ToString(dtAdd.Rows[j]["ITEM_NM"]);
                        sComSubCd1 = Convert.ToString(dtAdd.Rows[j]["GB_CD"]);
                        sComSubCd2 = Convert.ToString(dtAdd.Rows[j]["DET_CD"]).Substring(1, 1);
                        sComSubNm3 = Convert.ToString(dtAdd.Rows[j]["ITEM_NM"]);
                        sRemark = Convert.ToString(dtAdd.Rows[j]["RMK"]);
                        sUseYn = Convert.ToString(dtAdd.Rows[j]["USE_YN"]);
                        sSortSeq = Convert.ToString(dtAdd.Rows[j]["SORT_SEQ"]);

                        strSql.AppendLine(" ");
                        strSql.AppendLine(" INSERT INTO COM_BASE_CD");
                        strSql.AppendLine(" (CD_GB, COM_CD, COM_SUB_CD1, COM_SUB_CD2, COM_SUB_CD3, CD_GB_NM, COM_NM");
                        strSql.AppendLine(" , COM_SUB_NM1, COM_SUB_NM2, COM_SUB_NM3, SORT_SEQ, USE_YN, MFY_DT, MFY_ID, REMARK)"); 
                        strSql.AppendLine(" SELECT '0001'"); 
                        strSql.AppendLine("      , CONCAT('" + sComSubCd1 + "', '" + sComSubCd2 + "', RIGHT(CONCAT('00', CONVERT(MAX(COM_SUB_CD3) + 1, CHAR)), 2))"); 
                        strSql.AppendLine("      , '" + sComSubCd1 + "'"); 
                        strSql.AppendLine("      , '" + sComSubCd2 + "'"); 
                        strSql.AppendLine("      , RIGHT(CONCAT('00', CONVERT(MAX(COM_SUB_CD3) + 1, CHAR)), 2)"); 
                        strSql.AppendLine("      , '" + sCdGbNm + "'"); 
                        strSql.AppendLine("      , '" + sComNm + "'"); 
                        strSql.AppendLine("      , (SELECT COM_NM FROM COM_BASE_CD WHERE CD_GB = '0002' AND COM_CD = '" + sComSubCd1 + "' )"); 
                        strSql.AppendLine("      , (SELECT COM_NM FROM COM_BASE_CD WHERE CD_GB = '0003' AND COM_CD = CONCAT('" + sComSubCd1 + "', '" + sComSubCd2 + "') )"); 
                        strSql.AppendLine("      , '" + sComSubNm3 + "'"); 
                        strSql.AppendLine("      , '" + sSortSeq + "'"); 
                        strSql.AppendLine("      , '" + sUseYn + "'"); 
                        strSql.AppendLine("      , NOW()"); 
                        strSql.AppendLine("      , '" + sId + "'"); 
                        strSql.AppendLine("      , '" + sRemark + "'"); 
                        strSql.AppendLine("   FROM COM_BASE_CD"); 
                        strSql.AppendLine("  WHERE CD_GB = '0001'"); 
                        strSql.AppendLine("    AND COM_SUB_CD1 = '" + sComSubCd1 + "'"); 
                        strSql.AppendLine("    AND COM_SUB_CD2 = '" + sComSubCd2 + "'"); 

                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = strSql.ToString();
                        cmd.ExecuteNonQuery();
                    }
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            GridViewRetr.AddNewRow();
            GridViewRetr.ShowEditForm();
        }

        private void MesGradeCdDev_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Dispose();
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
                
            }
        }

        private void GridRetr_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }
    }
}