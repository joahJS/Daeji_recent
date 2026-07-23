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

namespace AccAdm
{
    public partial class QimClaimDev : DevExpress.XtraEditors.XtraForm
    {
        public QimClaimDev()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void QimClaimDev_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DateEditFrom.EditValue = DateTime.Now.ToString("yyyy-MM-dd");
            DateEditTo.EditValue = DateTime.Now.ToString("yyyy-MM-dd");

            DataTable dtLkupITCodeRetr = GetLookUpData("1", "1", "", "Y");
            DataTable dtLkupITCode = GetLookUpData("1", "", "", "Y");
            DataTable dtLkupStat = GetLookUpData("2", "3", "", "Y");

            ComLib.ComGrid.SetLookUpEdit(LkupITCodeRetr, dtLkupITCodeRetr, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupITCode, dtLkupITCode, "CD", "NM", "y");
            ComLib.ComGrid.SetLookUpEdit(LkupStat, dtLkupStat, "CD", "NM", "Y");
            
            RepositoryItemGridLookUpEdit iTCodeLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(iTCodeLkup, dtLkupITCodeRetr, GridClaim, GridColITCode, "CD", "NM", "");

            RepositoryItemGridLookUpEdit statLkup = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(statLkup, dtLkupStat, GridClaim, GridColCStat, "CD", "NM", "");
            
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            arrGrdView = new GridView[] { GridViewClaim };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
            BtnRetr_Click(null, null);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
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
            string sFrom = DateEditFrom.EditValue.ToString().Substring(0, 10);
            string sTo = DateEditTo.EditValue.ToString().Substring(0, 10);
            string sITCodeRetr = LkupITCodeRetr.EditValue.ToString();

            StringBuilder strSql = new StringBuilder();
            Cursor = Cursors.WaitCursor;

            strSql.AppendLine(" SELECT A.CLAIM_NO ");
            strSql.AppendLine("      , A.I_USER");
            strSql.AppendLine("      , A.T_DATE ");
            strSql.AppendLine("      , A.IT_CODE ");
            strSql.AppendLine("      , A.CONTENT");
            strSql.AppendLine("      , A.I_BIGO ");
            strSql.AppendLine("      , A.C_STAT");
            strSql.AppendLine("      , A.CLAIM_RQ ");
            strSql.AppendLine("   FROM QIM_CLAIM A ");
            strSql.AppendLine("  WHERE A.T_DATE >= '" + sFrom + "' ");
            strSql.AppendLine("    AND A.T_DATE <= '" + sTo + "' ");
            strSql.AppendLine("    AND (('' = '" + sITCodeRetr + "') OR (('' <> '" + sITCodeRetr + "') AND A.IT_CODE = '" + sITCodeRetr + "'))");
            strSql.AppendLine("  ORDER BY A.CLAIM_NO ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridClaim.DataSource = dt;

            Cursor = Cursors.Default;
        }
        private void LkupITCodeRetr_EditValueChanged(object sender, EventArgs e)
        {
            GetGridRetr();
        }

        private void LkupStat_EditvalueChanged(object sender, EventArgs e)
        {
            GetGridRetr();
        }

        private void GridViewClaim_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            TxtClaimNo.Text = GridViewClaim.GetFocusedRowCellValue("CLAIM_NO") == null ? "" : GridViewClaim.GetFocusedRowCellValue("CLAIM_NO").ToString();
            TxtIUser.Text = GridViewClaim.GetFocusedRowCellValue("I_USER") == null ? "" : GridViewClaim.GetFocusedRowCellValue("I_USER").ToString();
            DateEditRequest.EditValue = GridViewClaim.GetFocusedRowCellValue("T_DATE") == null ? "" : GridViewClaim.GetFocusedRowCellValue("T_DATE").ToString();
            LkupITCode.EditValue = GridViewClaim.GetFocusedRowCellValue("IT_CODE") == null ? "" : GridViewClaim.GetFocusedRowCellValue("IT_CODE").ToString();
            MemoContent.Text = GridViewClaim.GetFocusedRowCellValue("CONTENT") == null ? "" : GridViewClaim.GetFocusedRowCellValue("CONTENT").ToString();
            TxtBigo.Text = GridViewClaim.GetFocusedRowCellValue("I_BIGO") == null ? "" : GridViewClaim.GetFocusedRowCellValue("I_BIGO").ToString();
            LkupStat.EditValue = GridViewClaim.GetFocusedRowCellValue("C_STAT") == null ? "" : GridViewClaim.GetFocusedRowCellValue("C_STAT").ToString();
            TxtClaimRq.Text = GridViewClaim.GetFocusedRowCellValue("CLAIM_RQ") == null ? "" : GridViewClaim.GetFocusedRowCellValue("CLAIM_RQ").ToString();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            ClearAllForm(layoutControl1);
            TxtIUser.Focus();
        }

        public void ClearAllForm(Control Ctrl)
        {
            if (Ctrl.HasChildren)
            {
                foreach (Control ctrl in Ctrl.Controls)
                {
                    if (ctrl is DevExpress.XtraEditors.TextEdit)
                    {
                        if ((((DevExpress.XtraEditors.BaseEdit)ctrl).Name.Equals(LkupITCodeRetr.ToString())))
                        {
                            return;
                        }
                        else
                        {
                            (ctrl as DevExpress.XtraEditors.TextEdit).ResetText();
                        }
                    }

                    if (ctrl is DevExpress.XtraEditors.LookUpEdit)
                        (ctrl as DevExpress.XtraEditors.LookUpEdit).EditValue = "";

                    if (ctrl is DevExpress.XtraEditors.DateEdit)
                        (ctrl as DevExpress.XtraEditors.DateEdit).EditValue = DateTime.Now.ToString("yyyy-MM-dd");

                    if (ctrl is DevExpress.XtraEditors.ComboBoxEdit)
                        (ctrl as DevExpress.XtraEditors.ComboBoxEdit).ResetText();

                    if (ctrl is DevExpress.XtraEditors.RadioGroup)
                        (ctrl as DevExpress.XtraEditors.RadioGroup).SelectedIndex = 1;

                    if (ctrl.HasChildren)
                        ClearAllForm(ctrl);//Recursive
                }
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
            else if (sNullYn.Equals("3"))
            {
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }


            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = '0001'");

            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE A.CD_GB = 'C_STAT'");
            }
            else
            {
                strSql.AppendLine(" SELECT A.ACC_CD AS CD");
                strSql.AppendLine("      , A.ACC_NM AS NM");
                strSql.AppendLine("   FROM ACC_ACC_CD A");
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
            if (!ClsFunc.CheckCRUDAuthority("UPDATE", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 저장 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            Cursor = Cursors.WaitCursor;
            string sClaimNo = TxtClaimNo.Text;
            string sIUser = TxtIUser.Text;
            string sRequest = DateEditRequest.EditValue.ToString();
            string sITCode = LkupITCode.EditValue.ToString();
            string sContent = MemoContent.Text;
            string sBigo = TxtBigo.Text;
            string sStat = LkupStat.EditValue.ToString();
            string sClaimRq = TxtClaimRq.Text;

            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;

            try
            {
                strSql.AppendLine("MERGE INTO QIM_CLAIM AS a         ");
                strSql.AppendLine("    USING(SELECT                  ");
                strSql.AppendLine("        CLAIM_NO = '" + sClaimNo + "'             ");
                strSql.AppendLine("        , I_USER = '" + sIUser + "'           ");
                strSql.AppendLine("        , T_DATE = '" + sRequest + "'   ");
                strSql.AppendLine("        , IT_CODE = '" + sITCode + "'        ");
                strSql.AppendLine("        , CONTENT = '" + sContent + "'            ");
                strSql.AppendLine("        , I_BIGO = '" + sBigo + "'             ");
                strSql.AppendLine("        , C_STAT = '" + sStat + "'            ");
                strSql.AppendLine("        , CLAIM_RQ = '" + sClaimRq + "') AS b     ");
                strSql.AppendLine("    ON a.CLAIM_NO = b.CLAIM_NO    ");
                strSql.AppendLine("    WHEN MATCHED THEN UPDATE SET  ");
                strSql.AppendLine("        I_USER = '" + sIUser + "'             ");
                strSql.AppendLine("        , T_DATE = '" + sRequest + "'   ");
                strSql.AppendLine("        , IT_CODE = '" + sITCode + "'        ");
                strSql.AppendLine("        , CONTENT = '" + sContent + "'            ");
                strSql.AppendLine("        , I_BIGO = '" + sBigo + "'             ");
                strSql.AppendLine("        , C_STAT = '" + sStat + "'            ");
                strSql.AppendLine("        , CLAIM_RQ = '" + sClaimRq + "'   ");
                strSql.AppendLine("    WHEN NOT MATCHED THEN INSERT( ");
                strSql.AppendLine("         I_USER                   ");
                strSql.AppendLine("        , T_DATE                  ");
                strSql.AppendLine("        , IT_CODE                 ");
                strSql.AppendLine("        , CONTENT                 ");
                strSql.AppendLine("        , I_BIGO                  ");
                strSql.AppendLine("        , C_STAT                  ");
                strSql.AppendLine("        , CLAIM_RQ)               ");
                strSql.AppendLine("    VALUES('" + sIUser + "'                   ");
                strSql.AppendLine("        , '" + sRequest + "'            ");
                strSql.AppendLine("        , '" + sITCode + "'                  ");
                strSql.AppendLine("        , '" + sContent + "'                      ");
                strSql.AppendLine("        , '" + sBigo + "'                      ");
                strSql.AppendLine("        , '" + sStat + "'                     ");
                strSql.AppendLine("        , '" + sClaimRq + "');                    ");

                /*
                strSql.AppendLine(" INSERT INTO QIM_CLAIM  ");
                strSql.AppendLine("           ( CLAIM_NO ");
                strSql.AppendLine("           , I_USER ");
                strSql.AppendLine("           , T_DATE");
                strSql.AppendLine("           , IT_CODE");
                strSql.AppendLine("           , CONTENT");
                strSql.AppendLine("           , I_BIGO");
                strSql.AppendLine("           , C_STAT");
                strSql.AppendLine("           , CLAIM_RQ");
                strSql.AppendLine("           ) ");
                if (sClaimNo != "") strSql.AppendLine("        VALUES(    '" + sClaimNo + "'");
                else strSql.AppendLine(" VALUES    ( null ");
                strSql.AppendLine("           , '" + sIUser + "' ");
                strSql.AppendLine("           , '" + sRequest + "' ");
                strSql.AppendLine("           , '" + sITCode + "' ");
                strSql.AppendLine("           , '" + sContent + "' ");
                strSql.AppendLine("           , '" + sBigo + "' ");
                strSql.AppendLine("           , '" + sStat + "' ");
                strSql.AppendLine("           , '" + sClaimRq + "' ");
                strSql.AppendLine("           ) ");
                strSql.AppendLine(" ON DUPLICATE KEY UPDATE I_USER    = '" + sIUser + "' ");
                strSql.AppendLine("                       , T_DATE     = '" + sRequest + "' ");
                strSql.AppendLine("                       , IT_CODE     = '" + sITCode + "' ");
                strSql.AppendLine("                       , CONTENT     = '" + sContent + "' ");
                strSql.AppendLine("                       , I_BIGO     = '" + sBigo + "' ");
                strSql.AppendLine("                       , C_STAT     = '" + sStat + "' ");
                strSql.AppendLine("                       , CLAIM_RQ     = '" + sClaimRq + "' ");
                */

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                BtnRetr_Click(null, null);
            }
            catch (Exception ex)
            {
                DBConn.dbTran.Rollback();
                DBConn.dbTran = null;
                MessageBox.Show(ex.Message);
            }

            Cursor = Cursors.Default;
        }

        private void QimClaimDev_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
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
                BtnExcel_Click(null, null);
            }
        }

        private void GridClaim_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewClaim_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void QimClaimDev_FormClosed(object sender, FormClosedEventArgs e)
        {
            
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
                string sFileNM = "클레임관리 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;

                    GridClaim.ExportToXls(FileName + ".xls");
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

        private void QimClaimDev_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void LkupITCodeRetr_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}