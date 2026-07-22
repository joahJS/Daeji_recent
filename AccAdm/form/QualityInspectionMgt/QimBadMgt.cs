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
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid;
using System.Data.SqlClient;

namespace AccAdm
{
    public partial class QimBadMgt : DevExpress.XtraEditors.XtraForm
    {
        public QimBadMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void QimBadMgt_Load(object sender, EventArgs e)
        {
            DataTable dtLkupCvnamRetr = GetLookUpData("3", "2", "", "N");
            DataTable dtLkupGradeRetr = GetLookUpData("1", "1", "", "Y");
            DataTable dtLkupCvNamId = GetLookUpData("2", "1", "", "Y");
            DataTable dtLkupGradeCd = GetLookUpData("1", "1", "", "Y");

            ComLib.ComGrid.SetLookUpEdit(LkupCvnamRetr, dtLkupCvnamRetr, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupGradeRetr, dtLkupGradeRetr, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupCvNamId, dtLkupCvNamId, "CD", "NM", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupGradeCd, dtLkupGradeCd, "CD", "NM", "Y");

            
            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            arrGrdView = new GridView[] { GridViewBadCheck };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void ClearFps()
        {
            GridViewBadCheck.FocusedRowChanged -= GridViewBadCheck_FocusedRowChanged;
            GridBadCheck.DataSource = null;
            GridViewBadCheck.FocusedRowChanged += GridViewBadCheck_FocusedRowChanged;

        }
        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;

            }
            string sCvnamRetr = LkupCvnamRetr.EditValue?.ToString();
            string sGradeRetr = LkupGradeRetr.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();
            Cursor = Cursors.WaitCursor;

            strSql.AppendLine(" SELECT A.MAKENO ");
            strSql.AppendLine("      , A.MAKENO_LM");
            strSql.AppendLine("      , A.M4_CARNO ");
            strSql.AppendLine("      , A.M4_CVNAM ");
            strSql.AppendLine("      , A.M4_CVNAM_IDTNO");
            strSql.AppendLine("      , A.M4_GRADE ");
            strSql.AppendLine("      , A.M4_GRADE_CD");
            strSql.AppendLine("      , A.M4_WGT ");
            strSql.AppendLine("      , A.M4_MINUS");
            strSql.AppendLine("      , A.M4_EVIDENCE ");
            strSql.AppendLine("      , A.ENT_DT");
            strSql.AppendLine("      , A.ENT_ID ");
            strSql.AppendLine("      , A.MFY_DT");
            strSql.AppendLine("      , A.MFY_ID ");
            strSql.AppendLine("   FROM MAKE_4 A ");
            strSql.AppendLine("  WHERE 1=1");
            if (sCvnamRetr != "****") { 
            strSql.AppendLine("    AND A.M4_CVNAM_IDTNO = '" + sCvnamRetr + "' ");
            }
            if(sGradeRetr != "****") { 
            strSql.AppendLine("    AND A.M4_GRADE_CD = '" + sGradeRetr + "' ");
            }
            strSql.AppendLine("  ORDER BY A.MAKENO ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            if (dt.Rows.Count > 0) GridBadCheck.DataSource = dt;

            Cursor = Cursors.Default;

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
            else if (sNullYn.Equals("2"))
            {
                strSql.AppendLine(" SELECT '****' AS CD");
                strSql.AppendLine("     , '전체' AS NM");
                //strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }


            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.J_SERIAL AS CD");
                strSql.AppendLine("      , A.GUBUN1 AS NM");
                strSql.AppendLine("      , A.J_SERIAL AS SEQ");
                strSql.AppendLine("   FROM JAJAE A");

            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.IDT_NO AS CD");
                strSql.AppendLine("      , A.DEALER_NM AS NM");
                strSql.AppendLine("      , A.IDT_NO AS SEQ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT A.IDT_NO AS CD");
                strSql.AppendLine("      , A.DEALER_NM AS NM");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
                strSql.AppendLine(" WHERE DEALER_GB = '매입' OR DEALER_GB = '매출' ");
            }
            else
            {

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

        private void GridViewBadCheck_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            TxtMakeNo.Text = GridViewBadCheck.GetFocusedRowCellValue("MAKENO").ToString();
            TxtMakeNoLm.Text = GridViewBadCheck.GetFocusedRowCellValue("MAKENO_LM").ToString();
            TxtCarNo.Text = GridViewBadCheck.GetFocusedRowCellValue("M4_CARNO").ToString();
            LkupCvNamId.EditValue = GridViewBadCheck.GetFocusedRowCellValue("M4_CVNAM_IDTNO").ToString();
            MemoEvidence.Text = GridViewBadCheck.GetFocusedRowCellValue("M4_EVIDENCE").ToString();
            LkupGradeCd.EditValue = GridViewBadCheck.GetFocusedRowCellValue("M4_GRADE_CD").ToString();
            TxtWgt.Text = GridViewBadCheck.GetFocusedRowCellValue("M4_WGT").ToString();
            TxtMinus.Text = GridViewBadCheck.GetFocusedRowCellValue("M4_MINUS").ToString();
            DateEditEntDt.EditValue = GridViewBadCheck.GetFocusedRowCellValue("ENT_DT").ToString();
            TxtEntId.Text = GridViewBadCheck.GetFocusedRowCellValue("ENT_ID").ToString();
            DateEditMfyDt.EditValue = GridViewBadCheck.GetFocusedRowCellValue("MFY_DT").ToString();
            TxtMfyId.Text = GridViewBadCheck.GetFocusedRowCellValue("MFY_ID").ToString();
        }

        private void BtnNew_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            StringBuilder strSql = new StringBuilder();

            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT MAKENO_LM ");
            strSql.AppendLine("   FROM MAKE_4");
            strSql.AppendLine("  ORDER BY MAKENO_LM DESC");

            DataTable dtt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());

            string result = dtt.Rows[0]["MAKENO_LM"].ToString();
            double sMakeNoLmm = double.Parse(result) + 1;

            TxtMakeNo.Text = "";
            TxtMakeNoLm.Text = sMakeNoLmm.ToString();
            TxtCarNo.Text = "";
            LkupCvNamId.EditValue = "";
            MemoEvidence.Text = "";
            LkupGradeCd.EditValue = "";
            TxtWgt.Text = "";
            TxtMinus.Text = "";
            DateEditEntDt.EditValue = DateTime.Now;
            TxtEntId.Text = "";
            DateEditMfyDt.EditValue = "";
            TxtMfyId.Text = "";

            TxtMakeNo.ReadOnly = false;
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
            string sMakeNo = TxtMakeNo.Text;
            string sMakeNoLm = TxtMakeNoLm.Text;
            string sCarNo = TxtCarNo.Text;
            string sCvNam = LkupCvNamId.Text;
            string sCvNamId = LkupCvNamId.EditValue.ToString();
            string sEvidence = MemoEvidence.Text;
            string sGradeCd = LkupGradeCd.EditValue.ToString();
            string sGrade = LkupGradeCd.Text;
            string sWgt = TxtWgt.Text;
            string sMinus = TxtMinus.Text;
            string sEntDt = DateEditEntDt.EditValue.ToString();
            string sEntId = TxtEntId.Text;
            string sMfyDt = DateEditMfyDt.EditValue.ToString();
            string sMfyId = TxtMfyId.Text;

            ClearFps();


            DBConn.dbTran = DBConn.dbCon.BeginTransaction(IsolationLevel.ReadCommitted);
            SqlCommand cmd = DBConn.dbCon.CreateCommand();
            cmd.Transaction = DBConn.dbTran;
            strSql.Clear();
            

            try
            {
                strSql.AppendLine(" INSERT INTO MAKE_4  ");
                strSql.AppendLine("           ( MAKENO ");
                strSql.AppendLine("           , MAKENO_LM ");
                strSql.AppendLine("           , M4_CARNO ");
                strSql.AppendLine("           , M4_CVNAM");
                strSql.AppendLine("           , M4_CVNAM_IDTNO");
                strSql.AppendLine("           , M4_GRADE");
                strSql.AppendLine("           , M4_GRADE_CD");
                strSql.AppendLine("           , M4_WGT");
                strSql.AppendLine("           , M4_MINUS");
                strSql.AppendLine("           , M4_EVIDENCE");
                strSql.AppendLine("           , ENT_DT");
                strSql.AppendLine("           , ENT_ID");
                strSql.AppendLine("           , MFY_DT");
                strSql.AppendLine("           , MFY_ID");
                strSql.AppendLine("           ) ");
                strSql.AppendLine("        VALUES(    '" + sMakeNo + "'");
                strSql.AppendLine("           , '" + sMakeNoLm + "' ");
                strSql.AppendLine("           , '" + sCarNo + "' ");
                strSql.AppendLine("           , '" + sCvNam + "' ");
                strSql.AppendLine("           , '" + sCvNamId + "' ");
                strSql.AppendLine("           , '" + sGrade + "' ");
                strSql.AppendLine("           , '" + sGradeCd + "' ");
                strSql.AppendLine("           , '" + sWgt + "' ");
                strSql.AppendLine("           , '" + sMinus + "' ");
                strSql.AppendLine("           , '" + sEvidence + "' ");
                strSql.AppendLine("           , '" + sEntDt + "' ");
                strSql.AppendLine("           , '" + sEntId + "' ");
                strSql.AppendLine("           , '" + sMfyDt + "' ");
                strSql.AppendLine("           , '" + sMfyId + "' ");
                strSql.AppendLine("           ) ");
                strSql.AppendLine(" ON DUPLICATE KEY UPDATE M4_CARNO    = '" + sCarNo + "' ");
                strSql.AppendLine("                       , M4_CVNAM     = '" + sCvNam + "' ");
                strSql.AppendLine("                       , M4_CVNAM_IDTNO     = '" + sCvNamId + "' ");
                strSql.AppendLine("                       , M4_GRADE     = '" + sGrade + "' ");
                strSql.AppendLine("                       , M4_GRADE_CD     = '" + sGradeCd + "' ");
                strSql.AppendLine("                       , M4_WGT     = '" + sWgt + "' ");
                strSql.AppendLine("                       , M4_MINUS     = '" + sMinus + "' ");
                strSql.AppendLine("                       , M4_EVIDENCE     = '" + sEvidence + "' ");
                strSql.AppendLine("                       , ENT_DT     = '" + sEntDt + "' ");
                strSql.AppendLine("                       , ENT_ID     = '" + sEntId + "' ");
                strSql.AppendLine("                       , MFY_DT     = '" + sMfyDt + "' ");
                strSql.AppendLine("                       , MFY_ID     = '" + sMfyId + "' ");

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strSql.ToString();
                cmd.ExecuteNonQuery();

                DBConn.dbTran.Commit();
                DBConn.dbTran = null;
                MessageBox.Show("저장을 완료했습니다.");

                TxtMakeNo.ReadOnly = true;
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

        private void QimBadMgt_KeyDown(object sender, KeyEventArgs e)
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

            }
            else if (e.KeyCode == Keys.F5)
            {
                BtnRetr_Click(null, null);
            }
            else if (e.KeyCode == Keys.F8)
            {
                
            }
        }

        private void GridBadCheck_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void GridViewBadCheck_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void QimBadMgt_FormClosed(object sender, FormClosedEventArgs e)
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
                string sFileNM = "검수불량 리스트";
                string sFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                fileDlg.InitialDirectory = sFolderPath;
                fileDlg.FileName = sFileNM;
                fileDlg.Filter = "Excel files (*.xls or .xlsx)|.xls;*.xlsx";

                if (fileDlg.ShowDialog() == DialogResult.OK)
                {
                    FileName = fileDlg.FileName;

                    GridBadCheck.ExportToXls(FileName + ".xls");
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

        private void QimBadMgt_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }
    }
}