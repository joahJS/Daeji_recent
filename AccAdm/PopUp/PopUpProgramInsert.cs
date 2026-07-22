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
    public partial class PopUpProgramInsert : DevExpress.XtraEditors.XtraForm
    {
        public PopUpProgramInsert()
        {
            InitializeComponent();
        }

        public DataRow drPgm { get; set; }

        private void PopUpProgramInsert_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);
            DataTable dtGroup = GetLookUpData("1", "", "", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupGroup, dtGroup, "CD", "NM", "Y");
            if (drPgm != null)
            {
                TxtPgmId.EditValue = drPgm["PGMID"]?.ToString();
                TxtPgmNm.EditValue = drPgm["PGMNM"]?.ToString();
                LkupGroup.EditValue = drPgm["PGGRP"]?.ToString();
                RdgbWorkGb.EditValue = drPgm["PGGUB"]?.ToString();
                CboGrade.EditValue = drPgm["PGLVL"]?.ToString();
                TxtMenuTag.EditValue = drPgm["PGTAG"]?.ToString();
                TxtBigo.EditValue = drPgm["RK"]?.ToString();
                ChkUseYn.EditValue = drPgm["USEYN"]?.ToString();
            }
            TxtPgmNm.SelectAll();
            TxtPgmNm.Focus();
        }

        private DataTable GetLookUpData(string sGb, string sNullYn, string sParam, string sOther)
        {
            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");

            strSql.AppendLine("WITH ITEM_INFO AS (");

            if (sNullYn.Equals("1"))
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
                strSql.AppendLine("  WHERE A.CD_GB = 'Z1'");           
                
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , A.DEPT_CD AS SEQ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  WHERE DEPT_CD <> '0000'");
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
        private void BtnSave_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();

            string sPgmId = TxtPgmId.EditValue?.ToString();

            if (string.IsNullOrEmpty(sPgmId))
            {
                XtraMessageBox.Show("프로그램ID는 필수 입력사항 입니다.");
                return;
            }

            string sPgmNM = TxtPgmNm.EditValue?.ToString();
            string sPGGrp = LkupGroup.EditValue?.ToString();
            string sWorkGB = RdgbWorkGb.EditValue?.ToString();
            string sGrade = CboGrade.EditValue?.ToString();

            if(string.IsNullOrEmpty(sWorkGB) || string.IsNullOrEmpty(sGrade))
            {
                XtraMessageBox.Show("프로그램의 업무그룹 및 권한레벨을 설정해주세요.");
                return;
            }

            string sTag = TxtMenuTag.EditValue?.ToString();
            string sBigo = TxtBigo.EditValue?.ToString();
            //if (string.IsNullOrEmpty(sBigo))
            //{
            //    XtraMessageBox.Show("비고를 입력하여주세요.");
            //    return;
            //}

            string sUseYn = ChkUseYn.EditValue?.ToString();
            if (string.IsNullOrEmpty(sUseYn))
            {
                XtraMessageBox.Show("사용여부를 체크하여 주세요.");
                return;
            }

            try
            {
                DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
                SqlCommand cmd = DBConn.dbCon.CreateCommand();
                cmd.Transaction = DBConn.dbTran;

                #region LOG
                //string sLOG = string.Empty;
                //string sEDIT_KIND = string.Empty;

                //strSql.Clear();
                //strSql.AppendLine(" ");
                //strSql.AppendLine(" SELECT * ");
                //strSql.AppendLine("   FROM ZPGMLST ");
                //strSql.AppendLine("  WHERE PGMID = '" + sPgmId + "' ");

                //DataTable dtPrv = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
                //if (dtPrv.Rows.Count > 0)
                //{
                //    sLOG = string.Concat("[프로그램관리]프로그램ID : ", dtPrv.Rows[0]["PGMID"],
                //        ", 프로그램명 : ", dtPrv.Rows[0]["PGMNM"], " ▶ ", sPgmNM,
                //        ", 그룹 : ", dtPrv.Rows[0]["PGGRP"], " ▶ ", sPGGrp,
                //        ", 프로그램레벨 : ", dtPrv.Rows[0]["PGLVL"], " ▶ ", sGrade,
                //        ", 사용여부 : ", dtPrv.Rows[0]["USEYN"], " ▶ ", sUseYn);
                //}
                //else
                //{
                //    sLOG = string.Format("[프로그램관리]" +
                //       "프로그램ID : {0}" +
                //       "프로그램명 : {1}" +
                //       "그룹 : {2}" +
                //       "레벨 : {3}" +
                //       "사용여부 : {4}"
                //       , sPgmId
                //       , sPgmNM
                //       , sPGGrp
                //       , sGrade
                //       , sUseYn);
                //}
                #endregion

                strSql.Clear();
                strSql.AppendLine(" ");
                strSql.AppendLine(" IF EXISTS ( SELECT * FROM ZPGMLST WHERE PGMID = '"+ sPgmId + "' )");
                strSql.AppendLine("    BEGIN ");
                strSql.AppendLine(" 	      UPDATE ZPGMLST ");
                strSql.AppendLine("             SET PGMNM = '"+ sPgmNM + "' ");
                strSql.AppendLine("               , PGGRP = '"+ sPGGrp + "' ");
                strSql.AppendLine("               , PGGUB = '"+ sWorkGB + "' ");
                strSql.AppendLine("               , PGTAG = '"+ sTag + "' ");
                strSql.AppendLine("               , PGLVL = '"+ sGrade + "' ");
                strSql.AppendLine("               , USEYN = '"+ sUseYn + "' ");
                strSql.AppendLine("               , RK    = '"+ TxtBigo.EditValue?.ToString().Trim() + "'                             ");
                strSql.AppendLine("               , MUSER = '"+ FmMainToolBar2.UserID + "' ");
                strSql.AppendLine("               , MDATE = GETDATE() ");
                strSql.AppendLine("           WHERE PGMID = '"+ sPgmId + "' ");
                strSql.AppendLine("      END ");
                strSql.AppendLine(" ELSE ");
                strSql.AppendLine("    BEGIN ");
                strSql.AppendLine("          INSERT INTO ZPGMLST  ");
                strSql.AppendLine("                    ( PGMID, PGMNM, PGGRP  ");
                strSql.AppendLine("                    , PGGUB, PGTAG, PGLVL  ");
                strSql.AppendLine("                    , USEYN, RK   , CUSER  ");
                strSql.AppendLine("                    , CDATE ) ");
                strSql.AppendLine("              VALUES( '"+ sPgmId + "','"+ sPgmNM + "', '"+ sPGGrp + "'  ");
                strSql.AppendLine("                    , '"+ sWorkGB + "','"+ sTag + "', '"+ sGrade + "'  ");
                strSql.AppendLine("                    , '"+ sUseYn + "','"+ TxtBigo.EditValue?.ToString().Trim() + "', '"+ FmMainToolBar2.UserID + "'  ");
                strSql.AppendLine("                    , GETDATE() )");
                strSql.AppendLine("      END ");
                
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();
                cmd.Parameters.Clear();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;

                XtraMessageBox.Show("저장을 완료했습니다.");

                UserMgt fm = this.Owner as UserMgt;
                fm._PRGID = sPgmId;

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

        private void BtnClear_Click(object sender, EventArgs e)
        {
            TxtPgmId.Text = null;
            TxtPgmNm.Text = null;
            LkupGroup.EditValue = null;
            RdgbWorkGb.EditValue = null;
            CboGrade.Text = null;
            TxtMenuTag.Text = null;
            TxtBigo.Text = null;
            ChkUseYn.EditValue = null;
        }
        private void BtnClose_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void PopUpProgramInsert_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                BtnClear_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                BtnSave_Click(null, null);
            }
        }

        private void LkupGroup_Leave(object sender, EventArgs e)
        {
            string sWorkGroup = LkupGroup.EditValue?.ToString();

            if (string.IsNullOrEmpty(sWorkGroup))
            {
                XtraMessageBox.Show("업무그룹을 지정하여 주세요.");
                return;
            }

            if (sWorkGroup.Equals("11"))
            {
                CboGrade.EditValue = "0";
            }
            else if (sWorkGroup.Equals("77"))
            {
                CboGrade.EditValue = "1";
            }
            else if (sWorkGroup.Equals("66"))
            {
                CboGrade.EditValue = "5";
            }
            else
            {
                CboGrade.EditValue = "9";
            }
        }

        private void TxtBigo_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                BtnSave.Focus();
            }
        }

        private void LkupGroup_EditValueChanged(object sender, EventArgs e)
        {
            string sWorkGroup = LkupGroup.EditValue?.ToString();

            if (string.IsNullOrEmpty(sWorkGroup))
            {
                XtraMessageBox.Show("업무그룹을 지정하여 주세요.");
                return;
            }

            if (sWorkGroup.Equals("11"))
            {
                CboGrade.EditValue = "0";
            }
            else if (sWorkGroup.Equals("77"))
            {
                CboGrade.EditValue = "1";
            }
            else if (sWorkGroup.Equals("66"))
            {
                CboGrade.EditValue = "5";
            }
            else
            {
                CboGrade.EditValue = "9";
            }
        }
    }
}