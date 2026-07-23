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
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid;

namespace AccAdm
{
    public partial class EquipCdHistoryMgt : DevExpress.XtraEditors.XtraForm
    {
        public EquipCdHistoryMgt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void EquipCdHistoryMgt_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DateEditFrom.EditValue = DateTime.Today;
            DateEditTo.EditValue = DateTime.Today;

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);

            DataTable dtEquip = GetLookUpData("1", "Y", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupEquipCd, dtEquip, "CD", "NM", "");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupEquipCd, dtEquip, GridRetr, GridColMgNo, "CD", "NM", "");
            DataTable dtEmpId = GetLookUpData("2", "Y", "Y");
            ComLib.ComGrid.SetGridLookUpEdit(RepoGLkupEmpId, dtEmpId, GridRetr, GridColMgUser, "CD", "NM", "");

            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);

            BtnRetr_Click(null, null);
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
                strSql.AppendLine(" SELECT A.EMP_ID AS CD");
                strSql.AppendLine("      , A.EMP_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY EMP_ID) AS SEQ ");
                strSql.AppendLine("   FROM HR_EMP_BASIS A");
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

        private void BtnRetr_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("READ", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 읽기 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sYmdFrom = DateEditFrom.EditValue?.ToString().Substring(0, 10);
            string sYmdTo = DateEditTo.EditValue?.ToString().Substring(0, 10);
            string sEquipCd = LkupEquipCd.EditValue?.ToString();

            Cursor = Cursors.WaitCursor;

            StringBuilder strSql = new StringBuilder();
            
            strSql.AppendLine(" WITH INFO AS ( ");
            strSql.AppendLine("         SELECT MG_NO  ");
            strSql.AppendLine("              , MG_HIS_SEQ  ");
            strSql.AppendLine("              , OCCUR_DT  ");
            strSql.AppendLine("              , MG_DESC  ");
            strSql.AppendLine("              , MG_COST  ");
            strSql.AppendLine("              , MG_USER  ");
            strSql.AppendLine("              , MAKENO  ");
            strSql.AppendLine("              , MAKENO_LN  ");
            strSql.AppendLine("              , LN_ESEQ  ");
            strSql.AppendLine("              , 'N' AS MAKENO_YN  ");
            strSql.AppendLine("           FROM EQUIP_CD_HISTORY  ");
            strSql.AppendLine("          WHERE MAKENO IS NULL  ");
            strSql.AppendLine("            AND OCCUR_DT >= '" + sYmdFrom + "'  ");
            strSql.AppendLine("            AND OCCUR_DT <= '" + sYmdTo + "'  ");
            if (!string.IsNullOrEmpty(sEquipCd)) strSql.AppendLine("            AND MG_NO = '" + sEquipCd + "'  ");
            strSql.AppendLine("          UNION  ");
            strSql.AppendLine("         SELECT MG_NO  ");
            strSql.AppendLine("              , MG_HIS_SEQ  ");
            strSql.AppendLine("              , OCCUR_DT  ");
            strSql.AppendLine("              , MG_DESC  ");
            strSql.AppendLine("              , SUM(convert(integer,MG_COST)) AS MG_COST ");
            strSql.AppendLine("              , MG_USER ");
            strSql.AppendLine("              , MAKENO  ");
            strSql.AppendLine("              , MAKENO_LN ");
            strSql.AppendLine("              , LN_ESEQ  ");
            strSql.AppendLine("              , 'Y' AS MAKENO_YN  ");
            strSql.AppendLine("           FROM EQUIP_CD_HISTORY  ");
            strSql.AppendLine("          WHERE MAKENO IS NOT NULL ");
            strSql.AppendLine("            AND OCCUR_DT >= '" + sYmdFrom + "'  ");
            strSql.AppendLine("            AND OCCUR_DT <= '" + sYmdTo + "'   ");
            if (!string.IsNullOrEmpty(sEquipCd)) strSql.AppendLine("            AND MG_NO = '" + sEquipCd + "'  ");
            strSql.AppendLine("          GROUP BY MAKENO, MAKENO_LN, MG_NO, MG_HIS_SEQ, OCCUR_DT, MG_DESC,MG_USER, LN_ESEQ");
            strSql.AppendLine("         ) SELECT MG_NO, MG_HIS_SEQ, OCCUR_DT  ");
            strSql.AppendLine("                , MG_DESC, MG_COST, MG_USER  ");
            strSql.AppendLine("                , MAKENO, MAKENO_LN, LN_ESEQ  ");
            strSql.AppendLine("                , MAKENO_YN  ");
            strSql.AppendLine("             FROM INFO ");
            strSql.AppendLine("            ORDER BY OCCUR_DT, MG_NO, MG_HIS_SEQ ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;

            Cursor = Cursors.Default;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ClsFunc.CheckCRUDAuthority("ADD", rowUserInfo))
            {
                XtraMessageBox.Show("해당 사용자에 대하여 추가 권한이 없습니다. \r\n 관리자에게 문의하세요.");
                return;
            }

            string sUserId = FmMainToolBar2.drUser["USRCD"]?.ToString();
            
            string sEmpId = GetUserEmpID(sUserId);
            if (string.IsNullOrEmpty(sEmpId))
            {
                XtraMessageBox.Show("해당 사용자의 사원번호가 존재하지 않습니다. \r\n 접속코드 : " + sUserId + "");
                return;
            }

            PopupEquipCdHisMgt frm = new PopupEquipCdHisMgt();
            frm.AddModifyGB = "ADD";
            frm.EmpId = sEmpId;

            if(frm.ShowDialog() == DialogResult.OK)
            {
                BtnRetr_Click(null, null);
            }
        }

        private void GridViewRetr_RowClick(object sender, DevExpress.XtraGrid.Views.Grid.RowClickEventArgs e)
        {
            if(e.Clicks == 2)
            {
                string sYN = GridViewRetr.GetFocusedRowCellValue("MAKENO_YN")?.ToString();

                if (!sYN.Equals("Y"))
                {
                    string sUserId = FmMainToolBar2.drUser["USRCD"]?.ToString();

                    string sEmpId = GridViewRetr.GetFocusedRowCellValue("MG_USER")?.ToString();

                    PopupEquipCdHisMgt frm = new PopupEquipCdHisMgt();
                    frm.AddModifyGB = "MODIFY";
                    frm.EmpId = sEmpId;
                    frm.DrHisInfo = GridViewRetr.GetFocusedDataRow();
                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        BtnRetr_Click(null, null);
                    }
                }
                else if (sYN.Equals("Y"))
                {
                    string sMakeNo = GridViewRetr.GetFocusedRowCellValue("MAKENO")?.ToString();
                    string sMakeNoLn = GridViewRetr.GetFocusedRowCellValue("MAKENO_LN")?.ToString();
                    string sDate = GridViewRetr.GetFocusedRowCellValue("OCCUR_DT")?.ToString();
                    string sEmpId = GridViewRetr.GetFocusedRowCellValue("MG_USER")?.ToString();

                    ProdCostAdder frm = new ProdCostAdder();
                    frm.sMakeNo = sMakeNo;
                    frm.sMakeNoLn = sMakeNoLn;
                    frm.ProcessDate = sDate;
                    frm.sEmpID = sEmpId;

                    if(frm.ShowDialog() == DialogResult.OK)
                    {
                        BtnRetr_Click(null, null);
                    }
                }
            }
        }

        private string GetUserEmpID(string UsrCd) //접속ID의 인사번호 취득
        {
            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT ISNULL(INSANO, '') AS EMP_ID ");
            strSql.AppendLine("   FROM ZUSRLST ");
            strSql.AppendLine("  WHERE USRCD = '" + UsrCd + "' ");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            string sEmpId = dt.Rows[0]["EMP_ID"]?.ToString();

            return sEmpId;
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void EquipCdHistoryMgt_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            return;
        }

        private void EquipCdHistoryMgt_KeyDown(object sender, KeyEventArgs e)
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
                //BtnSave_Click(null, null);
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

        private void GridViewRetr_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            if (e.RowHandle % 2 == 0)
            {
                e.Appearance.BackColor = SystemColors.GradientInactiveCaption;
            }
        }

        private void EquipCdHistoryMgt_FormClosed(object sender, FormClosedEventArgs e)
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
                string sFileNM = "설비이력 리스트";
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

        private void EquipCdHistoryMgt_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void LkupEquipCd_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}