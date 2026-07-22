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
    public partial class PopupEquipCdHisMgt : DevExpress.XtraEditors.XtraForm
    {
        public PopupEquipCdHisMgt()
        {
            InitializeComponent();
        }

        public string AddModifyGB;
        public DataRow DrHisInfo;
        public string EmpId;
        private void PopupEquipCdHisMgt_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DataTable dtEquip = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEquipCd, dtEquip, "CD", "NM", "");

            if (AddModifyGB.Equals("ADD"))
            {
                ClearComponents();
                BtnCtnsAdd.Enabled = true;
            }
            else if (AddModifyGB.Equals("MODIFY"))
            {
                string sMgNo = DrHisInfo["MG_NO"]?.ToString();
                string sOccurDt = DrHisInfo["OCCUR_DT"]?.ToString();
                string sMgCost = DrHisInfo["MG_COST"]?.ToString();
                string sDesc = DrHisInfo["MG_DESC"]?.ToString();
                string sHisSeq = DrHisInfo["MG_HIS_SEQ"]?.ToString();

                LkupEquipCd.EditValue = sMgNo;
                DateEditOccur.EditValue = sOccurDt;
                TxtCost.EditValue = sMgCost;
                MemoDesc.EditValue = sDesc;
                TxtHisSeq.EditValue = sHisSeq;

                LkupEquipCd.Enabled = false;
                BtnCtnsAdd.Enabled = false;
            }
        }

        private void ClearComponents()
        {
            DateEditOccur.EditValue = DateTime.Today;
            LkupEquipCd.EditValue = null;
            TxtHisSeq.EditValue = string.Empty;
            TxtCost.EditValue = string.Empty;
            MemoDesc.EditValue = string.Empty;
        }

        #region[LookupEdit Settings]
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
                strSql.AppendLine(" SELECT A.MG_NO AS CD");
                strSql.AppendLine("      , A.EQUIP_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY MG_NO) AS SEQ ");
                strSql.AppendLine("   FROM EQUIP_CD A");
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
        #endregion[LookupEdit Settings]

        private void BtnCtnsAdd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(LkupEquipCd.EditValue?.ToString()))
            {
                XtraMessageBox.Show("설비를 선택하세요.");
                return;
            }
            else if (string.IsNullOrEmpty(DateEditOccur.EditValue?.ToString()))
            {
                XtraMessageBox.Show("발생일자를 선택하세요.");
                return;
            }
            else if (string.IsNullOrEmpty(TxtCost.EditValue?.ToString()))
            {
                XtraMessageBox.Show("비용을 선택하세요.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sEquipCd = LkupEquipCd.EditValue?.ToString();
                string sOccurDt = DateEditOccur.EditValue?.ToString().Substring(0, 10);
                string sCost = TxtCost.EditValue?.ToString();
                string sDesc = MemoDesc.EditValue?.ToString();
                double dHisSeq = 0;

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" SELECT ISNULL(MAX(MG_HIS_SEQ), 0) + 1 AS MAX_VALUE ");
                strSql.AppendLine("   FROM EQUIP_CD_HISTORY ");
                strSql.AppendLine("  WHERE MG_NO = '" + sEquipCd + "' ");

                DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                if (dtChk.Rows.Count > 0)
                {
                    dHisSeq = Convert.ToDouble(dtChk.Rows[0]["MAX_VALUE"]);
                }
                else
                {
                    dHisSeq = 1;
                }

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" INSERT INTO EQUIP_CD_HISTORY ");
                strSql.AppendLine("           ( ");
                strSql.AppendLine("             MG_NO ");
                strSql.AppendLine("           , OCCUR_DT ");
                strSql.AppendLine("           , MG_DESC ");
                strSql.AppendLine("           , MG_COST ");
                strSql.AppendLine("           , MG_USER ");
                strSql.AppendLine("           ) ");
                strSql.AppendLine("      VALUES ");
                strSql.AppendLine("           ( ");
                strSql.AppendLine("             '" + sEquipCd + "' ");
                strSql.AppendLine("           , '" + sOccurDt + "' ");
                strSql.AppendLine("           , '" + sDesc + "' ");
                strSql.AppendLine("           , '" + sCost + "' ");
                strSql.AppendLine("           , '" + EmpId + "' ");
                strSql.AppendLine("           ) ");
                
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                Cursor = Cursors.Default;

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                ClearComponents();
                TxtCost.Focus();
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(LkupEquipCd.EditValue?.ToString()))
            {
                XtraMessageBox.Show("설비를 선택하세요.");
                return;
            }
            else if (string.IsNullOrEmpty(DateEditOccur.EditValue?.ToString()))
            {
                XtraMessageBox.Show("발생일자를 선택하세요.");
                return;
            }
            else if (string.IsNullOrEmpty(TxtCost.EditValue?.ToString()))
            {
                XtraMessageBox.Show("비용을 선택하세요.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sEquipCd = LkupEquipCd.EditValue?.ToString();
                string sOccurDt = DateEditOccur.EditValue?.ToString().Substring(0, 10);
                string sCost = TxtCost.EditValue?.ToString();
                string sDesc = MemoDesc.EditValue?.ToString();
                
                double dHisSeq = 0;
                
                if (AddModifyGB.Equals("ADD"))
                {
                    strSql.Clear();
                    strSql.AppendLine(" ");
                    strSql.AppendLine(" SELECT ISNULL(MAX(MG_HIS_SEQ), 0) + 1 AS MAX_VALUE ");
                    strSql.AppendLine("   FROM EQUIP_CD_HISTORY ");
                    strSql.AppendLine("  WHERE MG_NO = '" + sEquipCd + "' ");

                    DataTable dtChk = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

                    if (dtChk.Rows.Count > 0)
                    {
                        dHisSeq = Convert.ToDouble(dtChk.Rows[0]["MAX_VALUE"]);
                    }
                    else
                    {
                        dHisSeq = 1;
                    }
                }
                else if (AddModifyGB.Equals("MODIFY"))
                {
                    string sHisSeq = TxtHisSeq.EditValue?.ToString();
                    if (string.IsNullOrEmpty(sHisSeq))
                    {
                        dHisSeq = 1;
                    }
                    else
                    {
                        dHisSeq = Convert.ToDouble(sHisSeq);
                    }
                }

                strSql.Clear();
                strSql.AppendLine(" ");

                strSql.AppendLine("MERGE INTO EQUIP_CD_HISTORY AS a                         ");
                strSql.AppendLine("    USING(SELECT                                         ");
                strSql.AppendLine("                MG_NO = '" + sEquipCd + "'                               ");
                strSql.AppendLine("                , MG_HIS_SEQ =  " + dHisSeq + "                       ");
                strSql.AppendLine("                , OCCUR_DT = '" + sOccurDt + "'                ");
                strSql.AppendLine("                , MG_DESC = '" + sDesc + "'                           ");
                strSql.AppendLine("                , MG_COST = '" + sCost + "'                        ");
                strSql.AppendLine("                , MG_USER = '" + EmpId + "') AS b              ");
                strSql.AppendLine("    ON a.MG_NO = b.MG_NO AND a.MG_HIS_SEQ = b.MG_HIS_SEQ ");
                strSql.AppendLine("    WHEN MATCHED THEN UPDATE SET                         ");
                strSql.AppendLine("                MG_NO = '" + sEquipCd + "'                               ");
                strSql.AppendLine("                , OCCUR_DT = '" + sOccurDt + "'                ");
                strSql.AppendLine("                , MG_DESC = '" + sDesc + "'                           ");
                strSql.AppendLine("                , MG_COST = '" + sCost + "'                        ");
                strSql.AppendLine("                , MG_USER = '" + EmpId + "'                    ");
                strSql.AppendLine("    WHEN NOT MATCHED THEN INSERT(                        ");
                strSql.AppendLine("                MG_NO                                    ");
                strSql.AppendLine("               , OCCUR_DT                                ");
                strSql.AppendLine("               , MG_DESC                                 ");
                strSql.AppendLine("               , MG_COST                                 ");
                strSql.AppendLine("               , MG_USER)                                ");
                strSql.AppendLine("    VALUES(                                              ");
                strSql.AppendLine("                '" + sEquipCd + "'                                    ");
                strSql.AppendLine("               , '" + sOccurDt + "'                            ");
                strSql.AppendLine("               , '" + sDesc + "'                                      ");
                strSql.AppendLine("               , '" + sCost + "'                                   ");
                strSql.AppendLine("               , '" + EmpId + "');                             ");
               
                
                /*
                strSql.AppendLine(" INSERT INTO EQUIP_CD_HISTORY ");
                strSql.AppendLine("           ( ");
                strSql.AppendLine("             MG_NO ");
                strSql.AppendLine("           , MG_HIS_SEQ ");
                strSql.AppendLine("           , OCCUR_DT ");
                strSql.AppendLine("           , MG_DESC ");
                strSql.AppendLine("           , MG_COST ");
                strSql.AppendLine("           , MG_USER ");
                strSql.AppendLine("           ) ");
                strSql.AppendLine("      VALUES ");
                strSql.AppendLine("           ( ");
                strSql.AppendLine("             '" + sEquipCd + "' ");
                strSql.AppendLine("           ,  " + dHisSeq + " ");
                strSql.AppendLine("           , '" + sOccurDt + "' ");
                strSql.AppendLine("           , '" + sDesc + "' ");
                strSql.AppendLine("           , '" + sCost + "' ");
                strSql.AppendLine("           , '" + EmpId + "' ");
                strSql.AppendLine("           ) ");
                strSql.AppendLine("          ON DUPLICATE KEY UPDATE ");
                strSql.AppendLine("             OCCUR_DT = '" + sOccurDt + "' ");
                strSql.AppendLine("           , MG_DESC = '" + sDesc + "' ");
                strSql.AppendLine("           , MG_COST = '" + sCost + "' ");
                strSql.AppendLine("           , MG_USER = '" + EmpId + "' ");
                */

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                Cursor = Cursors.Default;

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장을 완료했습니다.");

                this.DialogResult = DialogResult.OK;
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;

                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
            
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void PopupEquipCdHisMgt_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void PopupEquipCdHisMgt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnCtnsAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
        }
    }
}