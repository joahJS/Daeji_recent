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
    public partial class AC01001F02 : DevExpress.XtraEditors.XtraForm
    {
        public AC01001F02()
        {
            InitializeComponent();
        }

        public string AddModifyGb { get; set; }
        public string AccCd { get; set; }

        public DataRow rowUserInfo { get; set; }
        private void AC01001F02_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DataTable dtAcrDr = GetLookUpData("1", "Y", "Y");
            DataTable dtAgubun = GetLookUpData("2", "Y", "Y");

            ComLib.ComGrid.SetLookUpEdit(LkupAcrDr, dtAcrDr, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupAgubun, dtAgubun, "CD", "NM", "Y");

            FmMainToolBar2._FontSetting.SetGridView(GridViewRmk);
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            if (AddModifyGb.Equals("MOD"))
            {
                if (string.IsNullOrEmpty(AccCd))
                {
                    XtraMessageBox.Show("계정코드가 존재하지 않습니다.");
                    DialogResult = DialogResult.Cancel;
                    return;
                }

                DataTable dt = GetAccountInfo(AccCd);

                if(dt.Rows.Count > 0)
                {
                    string sAccCd = dt.Rows[0]["ACCOD"]?.ToString();
                    string sAcNam = dt.Rows[0]["ACNAM"]?.ToString();
                    string sAcDsp = dt.Rows[0]["ACDSP"]?.ToString();
                    string sAGubn = dt.Rows[0]["AGUBN"]?.ToString();
                    string sAcrDr = dt.Rows[0]["ACRDR"]?.ToString();
                    string sAsmCd = dt.Rows[0]["ASMCD"]?.ToString();
                    string sAsmNm = dt.Rows[0]["ASMNM"]?.ToString();
                    string sChgYn = dt.Rows[0]["CHGYN"]?.ToString();
                    string sCkCod = dt.Rows[0]["CKCOD"]?.ToString();
                    string sAbgGn = dt.Rows[0]["ABGGN"]?.ToString();
                    string sAjnGn = dt.Rows[0]["AJNGN"]?.ToString();
                    string sUseYn = dt.Rows[0]["USEYN"]?.ToString();
                    string sRk = dt.Rows[0]["RK"]?.ToString();

                    BtnEditAccCd.EditValue = sAccCd;
                    TxtAccNm.EditValue = sAcNam;
                    TxtLookNm.EditValue = sAcDsp;
                    LkupAgubun.EditValue = sAGubn;
                    LkupAcrDr.EditValue = sAcrDr;
                    BtnEditAsmCd.EditValue = sAsmCd;
                    TxtAsmNm.EditValue = sAsmNm;
                    ChkChgYn.EditValue = sChgYn;
                    ChkDealerYn.EditValue = sCkCod;
                    RdgbAbgGn.EditValue = sAbgGn;
                    RdgbAjnGn.EditValue = sAjnGn;
                    ChkUseYn.EditValue = sUseYn;
                    MemoEditRemark.EditValue = sRk;

                    GridRmk.DataSource = GetAccountRemarkInfo(AccCd);

                    BtnCntnsAdd.Enabled = false;
                    BtnEditAccCd.Enabled = false;

                    string sChk = ChkChgYn.EditValue == null ? "N" : ChkChgYn.EditValue.ToString();
                    if (sChk.Equals("Y"))
                    {
                        TxtLookNm.Enabled = true;
                        TxtLookNm.SelectAll();
                    }
                    else
                    {
                        TxtLookNm.Enabled = false;
                        BtnEditAsmCd.SelectAll();
                    }
                }
            }
            else
            {
                TxtLookNm.Enabled = true;
                BtnCntnsAdd.Enabled = true;
                BtnEditAccCd.Enabled = true;

                ChkChgYn.EditValue = "N";
                ChkDealerYn.EditValue = "N";
                ChkUseYn.EditValue = "Y";
            }
        }

        private void AC01001F02_Shown(object sender, EventArgs e)
        {
            if (AddModifyGb.Equals("ADD"))
            {
                BtnEditAccCd.Focus();
            }
            else
            {
                TxtAccNm.Focus();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            SaveAccountInfo("SAVE");
        }
        
        private void BtnCntnsAdd_Click(object sender, EventArgs e)
        {
            SaveAccountInfo("NOTSAVE");
        }
        
        private void EditClear()
        {
            BtnEditAccCd.EditValue = null;
            TxtAccNm.EditValue = null;
            TxtLookNm.EditValue = null;
            BtnEditAsmCd.EditValue = null;
            TxtAsmNm.EditValue = null;
            LkupAgubun.EditValue = null;
            LkupAcrDr.EditValue = null;
            RdgbAbgGn.EditValue = null;
            RdgbAjnGn.EditValue = null;
            MemoEditRemark.EditValue = null;

            ChkChgYn.EditValue = "N";
            ChkDealerYn.EditValue = "N";
            ChkUseYn.EditValue = "Y";

            GridRmk.DataSource = GetAccountRemarkInfo("");
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
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
            
            if (sGb.Equals("1")) //차대구분
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC01001_02 '");
            }
            else if (sGb.Equals("2")) //계정성격
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD ");
                strSql.AppendLine("      , A.COM_NM AS NM ");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ ");
                strSql.AppendLine("   FROM COM_BASE_CD A ");
                strSql.AppendLine("  WHERE A.CD_GB = 'AC01001_03 '");
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

        #region[Execute By Query]

        private DataTable GetAccountInfo(string sAccCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT A.ACCOD   ");
            strSql.AppendLine("      , A.ACNAM   ");
            strSql.AppendLine("      , A.ACDSP   ");
            strSql.AppendLine("      , A.AGUBN   ");
            strSql.AppendLine("      , A.ACRDR   ");
            strSql.AppendLine("      , A.ASMCD   ");
            strSql.AppendLine("      , B.ACNAM AS ASMNM ");
            strSql.AppendLine("      , A.CHGYN   ");
            strSql.AppendLine("      , A.CKCOD   ");
            strSql.AppendLine("      , A.ABGGN   ");
            strSql.AppendLine("      , A.AJNGN   ");
            strSql.AppendLine("      , A.USEYN   ");
            strSql.AppendLine("      , A.RK   ");
            strSql.AppendLine("      , A.CUSER   ");
            strSql.AppendLine("      , A.CDATE   ");
            strSql.AppendLine("      , A.MUSER   ");
            strSql.AppendLine("      , A.MDATE   ");
            strSql.AppendLine("   FROM ACMSTF A  ");
            strSql.AppendLine("   LEFT OUTER JOIN ACMSTF B ");
            strSql.AppendLine("     ON A.ASMCD = B.ACCOD ");
            strSql.AppendLine("  WHERE 1 = 1");
            strSql.AppendLine("    AND A.ACCOD = '" + sAccCd + "' ");
            
            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            return dt;
        }

        private DataTable GetAccountRemarkInfo(string sAccCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT * ");
            strSql.AppendLine("   FROM ( ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC1, '') AS RESULT ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC2, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC3, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC4, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC5, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC6, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC7, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC8, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("           UNION ALL ");
            strSql.AppendLine("          SELECT ISNULL(A.HDEC9, '') ");
            strSql.AppendLine("            FROM ACMSTF A ");
            strSql.AppendLine("           WHERE ACCOD = '" + sAccCd + "' ");
            strSql.AppendLine("        ) AS X1 ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            //if(dt == null || dt.Rows.Count <= 0)
            //{
            //    DataTable newTable = new DataTable();

            //    newTable.Columns.Add("RESULT");

            //    for(int i = 0; i < 9; i++)
            //    {
            //        DataRow row = newTable.NewRow();

            //        newTable.Rows.Add(row);
            //    }

            //    dt = newTable;
            //}

            return dt;
        }

        private void SaveAccountInfo(string sSaveGb)
        {
            string sChkAccCd = BtnEditAccCd.EditValue?.ToString().Replace(" ", "");
            string sChkAcDsp = TxtLookNm.EditValue?.ToString().Replace(" ", "");

            if (string.IsNullOrEmpty(sChkAccCd))
            {
                XtraMessageBox.Show("계정코드를 입력하세요.");
                BtnEditAccCd.Focus();
                BtnEditAccCd.SelectAll();
                return;
            }

            if (string.IsNullOrEmpty(sChkAcDsp))
            {
                XtraMessageBox.Show("표시계정명을 입력하세요.");
                TxtLookNm.Focus();
                TxtLookNm.SelectAll();
                return;
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                string sAccCd = BtnEditAccCd.EditValue?.ToString();
                string sAcNam = TxtAccNm.EditValue?.ToString();
                string sAcDsp = TxtLookNm.EditValue?.ToString();

                if (AddModifyGb.Equals("ADD"))
                    sAcNam = sAcDsp.Replace(" ", "");

                string sAGubn = LkupAgubun.EditValue?.ToString();
                string sAcrDr = LkupAcrDr.EditValue?.ToString();
                string sAsmCd = BtnEditAsmCd.EditValue?.ToString();
                string sAsmNm = TxtAsmNm.EditValue?.ToString();
                string sChgYn = ChkChgYn.EditValue?.ToString();
                string sCkCod = ChkDealerYn.EditValue?.ToString();
                string sAbgGn = RdgbAbgGn.EditValue?.ToString();
                string sAjnGn = RdgbAjnGn.EditValue?.ToString();
                string sUseYn = ChkUseYn.EditValue?.ToString();
                string sRk = MemoEditRemark.EditValue?.ToString();
                string sId = FmMainToolBar2.UserID;

                StringBuilder strSql = new StringBuilder();
                
                strSql.AppendLine("MERGE INTO ACMSTF as a               ");
                strSql.AppendLine("USING(SELECT ACCOD = '" + sAccCd + "'");
                strSql.AppendLine("       , ACNAM = '" + sAcNam + "'");
                strSql.AppendLine("       , ACDSP = '" + sAcDsp + "'");
                strSql.AppendLine("       , AGUBN = '" + sAGubn + "'");
                strSql.AppendLine("       , ACRDR = '" + sAcrDr + "'");
                strSql.AppendLine("       , ASMCD = '" + sAsmCd + "'");
                strSql.AppendLine("       , CHGYN = '" + sChgYn + "'");
                strSql.AppendLine("       , CKCOD = '" + sCkCod + "'");
                strSql.AppendLine("       , ABGGN = '" + sAbgGn + "'");
                strSql.AppendLine("       , AJNGN = '" + sAjnGn + "'");
                strSql.AppendLine("       , USEYN = '" + sUseYn + "'");
                strSql.AppendLine("       , RK = '" + sRk + "'");
                strSql.AppendLine("       , MUSER = '" + sId + "'");
                strSql.AppendLine("       , MDATE = CONVERT(VARCHAR(20),getdate(),20)) AS b     ");
                strSql.AppendLine(" ON a.ACCOD = b.ACCOD               ");
                strSql.AppendLine(" when MATCHED  then UPDATE  set      ");
                strSql.AppendLine("         ACCOD = '" + sAccCd + "'");
                strSql.AppendLine("       , ACNAM = '" + sAcNam + "'");
                strSql.AppendLine("       , ACDSP = '" + sAcDsp + "'");
                strSql.AppendLine("       , AGUBN = '" + sAGubn + "'");
                strSql.AppendLine("       , ACRDR = '" + sAcrDr + "'");
                strSql.AppendLine("       , ASMCD = '" + sAsmCd + "'");
                strSql.AppendLine("       , CHGYN = '" + sChgYn + "'");
                strSql.AppendLine("       , CKCOD = '" + sCkCod + "'");
                strSql.AppendLine("       , ABGGN = '" + sAbgGn + "'");
                strSql.AppendLine("       , AJNGN = '" + sAjnGn + "'");
                strSql.AppendLine("       , USEYN = '" + sUseYn + "'");
                strSql.AppendLine("       , RK = '" + sRk + "'");
                strSql.AppendLine("       , MUSER = '" + sId + "'");
                strSql.AppendLine("       , MDATE = CONVERT(VARCHAR(20),getdate(),20)          ");
                strSql.AppendLine(" WHEN NOT MATCHED THEN INSERT(       ");
                strSql.AppendLine("           ACCOD                     ");
                strSql.AppendLine("        , ACNAM                      ");
                strSql.AppendLine("       , ACDSP                       ");
                strSql.AppendLine("       , AGUBN                       ");
                strSql.AppendLine("       , ACRDR                       ");
                strSql.AppendLine("       , ASMCD                       ");
                strSql.AppendLine("       , CHGYN                       ");
                strSql.AppendLine("       , CKCOD                       ");
                strSql.AppendLine("       , ABGGN                       ");
                strSql.AppendLine("       , AJNGN                       ");
                strSql.AppendLine("       , USEYN                       ");
                strSql.AppendLine("       , RK                          ");
                strSql.AppendLine("       , MUSER                       ");
                strSql.AppendLine("       , MDATE                       ");
                strSql.AppendLine(" )VALUES(                            ");
                strSql.AppendLine("        '" + sAccCd + "'             ");
                strSql.AppendLine("       ,'" + sAcNam + "'             ");
                strSql.AppendLine("       , '" + sAcDsp + "'            ");
                strSql.AppendLine("       , '" + sAGubn + "'            ");
                strSql.AppendLine("       , '" + sAcrDr + "'            ");
                strSql.AppendLine("       , '" + sAsmCd + "'            ");
                strSql.AppendLine("       , '" + sChgYn + "'            ");
                strSql.AppendLine("       , '" + sCkCod + "'            ");
                strSql.AppendLine("       , '" + sAbgGn + "'            ");
                strSql.AppendLine("       , '" + sAjnGn + "'            ");
                strSql.AppendLine("       , '" + sUseYn + "'            ");
                strSql.AppendLine("       , '" + sRk + "'               ");
                strSql.AppendLine("       , '" + sId + "'               ");
                strSql.AppendLine("       , CONVERT(VARCHAR(20),getdate(),20)                  ");
                strSql.AppendLine(" );                                  ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                if (GridViewRmk.RowCount > 0)
                {
                    string[] sRmkArr = new string[GridViewRmk.RowCount];

                    for (int i = 0; i < GridViewRmk.RowCount; i++)
                    {
                        sRmkArr[i] = GridViewRmk.GetRowCellValue(i, "RESULT")?.ToString();
                    }

                    strSql.Clear();
                    strSql.AppendLine(" UPDATE ACMSTF ");
                    strSql.AppendLine("    SET HDEC1 = '" + sRmkArr[0] + "' ");
                    strSql.AppendLine("      , HDEC2 = '" + sRmkArr[1] + "' ");
                    strSql.AppendLine("      , HDEC3 = '" + sRmkArr[2] + "' ");
                    strSql.AppendLine("      , HDEC4 = '" + sRmkArr[3] + "' ");
                    strSql.AppendLine("      , HDEC5 = '" + sRmkArr[4] + "' ");
                    strSql.AppendLine("      , HDEC6 = '" + sRmkArr[5] + "' ");
                    strSql.AppendLine("      , HDEC7 = '" + sRmkArr[6] + "' ");
                    strSql.AppendLine("      , HDEC8 = '" + sRmkArr[7] + "' ");
                    strSql.AppendLine("      , HDEC9 = '" + sRmkArr[8] + "' ");
                    strSql.AppendLine("      , MUSER = '" + sId + "' ");
                    strSql.AppendLine("      , MDATE = CONVERT(VARCHAR(20),getdate(),20) ");
                    strSql.AppendLine("  WHERE ACCOD = '" + sAccCd + "' ");

                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = strSql.ToString();
                    cmd.ExecuteNonQuery();
                }

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                XtraMessageBox.Show("저장이 완료되었습니다.");

                AC01001F01 pFrm = (AC01001F01)this.Owner;

                pFrm._ACCOD = sAccCd;
                pFrm.GetGridRetrData();

                if (sSaveGb.Equals("SAVE"))
                {
                    DialogResult = DialogResult.OK;
                }
                else if (sSaveGb.Equals("NOTSAVE"))
                {
                    BtnEditAccCd.Focus();
                    EditClear();
                }
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                XtraMessageBox.Show(ex.Message);
            }
        }

        #endregion[Execute By Query]

        #region[KEY_DOWN_EVENT]

        private void AC01001F02_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnCntnsAdd_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
        }

        #endregion[KEY_DOWN_EVENT]

        public DataRow DrAccInfo;
        private void BtnEditAsmCd_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            string sAsmNm = BtnEditAsmCd.EditValue?.ToString();

            AC01001F03 frm = new AC01001F03();
            frm.PAC01001F02 = this;
            frm.AccCd = sAsmNm;
            if (frm.ShowDialog(this) == DialogResult.OK)
            {
                if(DrAccInfo.ItemArray.Length > 0)
                {
                    string sAsmCd = DrAccInfo["ACCOD"].ToString();
                    string sAsmNmResult = DrAccInfo["ACNAM"].ToString();
                    BtnEditAsmCd.EditValue = sAsmCd;
                    TxtAsmNm.EditValue = sAsmNmResult;
                }
            }
        }

        private void BtnEditAccCd_EditValueChanged(object sender, EventArgs e)
        {
            if (AddModifyGb.Equals("MOD"))
                return;

            string sAccCd = BtnEditAccCd.EditValue?.ToString();
            if (string.IsNullOrEmpty(sAccCd))
            {
                return;
            }
            else
            {
                if (CheckValueDuplicated(sAccCd))
                {
                    XtraMessageBox.Show("해당 계정코드는 존재하는 코드입니다.\r\n다시 입력하세요.");
                    BtnEditAccCd.Focus();
                    BtnEditAccCd.SelectAll();
                    return;
                }
            }
        }

        private bool CheckValueDuplicated(string sAccCd)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" SELECT COUNT(*) AS VALUE_CNT ");
            strSql.AppendLine("   FROM ACMSTF A ");
            strSql.AppendLine("  WHERE A.ACCOD = '" + sAccCd + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if(dt.Rows.Count > 0)
            {
                double dCnt = Convert.ToDouble(dt.Rows[0]["VALUE_CNT"]);
                if (dCnt == 0)
                    return false;
                else
                    return true;
            }
            else
            {
                return true;
            }
        }

        #region[GridView Row's Stripe Pattern]
        
        private void GridViewRmk_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        #endregion[GridView Row's Stripe Pattern]

        private void MemoEditRemark_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                BtnSave.Focus();
            }
        }
    }
}