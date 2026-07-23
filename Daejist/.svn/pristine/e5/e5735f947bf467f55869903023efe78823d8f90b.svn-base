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
using DevExpress.XtraEditors.Repository;
using System.Diagnostics;
using DevExpress.XtraGrid.Views.Grid;

namespace AccAdm
{
    public partial class EquipConsumeInOutRpt : DevExpress.XtraEditors.XtraForm
    {
        public EquipConsumeInOutRpt()
        {
            InitializeComponent();
        }

        public DataRow rowUserInfo { get; set; }
        public GridView[] arrGrdView;
        private void EquipConsumeInOutRpt_Load(object sender, EventArgs e)
        {
            ComnEtcFunc.gp_SetColorFocused(layoutControl1);

            DateEditFrom.EditValue = DateTime.Now;
            DateEditTo.EditValue = DateTime.Now;

            DataTable dtMenuItemNm = GetLookUpData("4", "Y", "Y");
            ComLib.ComGrid.SetLookUpEdit(LkupMenuItemNm, dtMenuItemNm, "CD", "NM", "");

            DataTable dtInOutGb = GetLookUpData("2", "Y", "Y");
            RepositoryItemGridLookUpEdit repoInOutGb = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(repoInOutGb, dtInOutGb, GridRetr, GridColInOutGb, "CD", "NM", "");

            DataTable dtDealer = GetLookUpData("3", "Y", "Y");
            RepositoryItemGridLookUpEdit repoDealer = new RepositoryItemGridLookUpEdit();
            ComLib.ComGrid.SetGridLookUpEdit(repoDealer, dtDealer, GridRetr, GridColSuplDealer, "CD", "NM", "");

            arrGrdView = new GridView[] { GridViewRetr };
            ComnEtcFunc.SetLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);

            rowUserInfo = ClsFunc.GetAutorityInfo(FmMainToolBar2.dtUserAutInfo, this.Name);
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
                strSql.AppendLine(" SELECT '' AS CD");
                strSql.AppendLine("     , '공용' AS NM");
                strSql.AppendLine("     , 0 AS SEQ");
                strSql.AppendLine(" UNION ALL");
            }
            
            if (sGb.Equals("1"))
            {
                strSql.AppendLine(" SELECT A.DEPT_CD AS CD");
                strSql.AppendLine("      , A.DEPT_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEPT_CD) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEPT_CD A");
                strSql.AppendLine("  WHERE USE_YN = 'Y'");
            }
            else if (sGb.Equals("2"))
            {
                strSql.AppendLine(" SELECT A.COM_CD AS CD");
                strSql.AppendLine("      , A.COM_NM AS NM");
                strSql.AppendLine("      , A.SORT_SEQ AS SEQ");
                strSql.AppendLine("   FROM COM_BASE_CD A");
                strSql.AppendLine("  WHERE CD_GB = 'ITEM_INOUT_GB'");
            }
            else if (sGb.Equals("3"))
            {
                strSql.AppendLine(" SELECT IDT_NO AS CD ");
                strSql.AppendLine("      , DEALER_NM AS NM ");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY DEALER_NM) AS SEQ ");
                strSql.AppendLine("   FROM ACC_DEALER_CD A");
                strSql.AppendLine("  WHERE EOB_YN = 'N' ");
            }
            else if (sGb.Equals("4"))
            {
                strSql.AppendLine(" SELECT CONVERT(VARCHAR,A.CON_ITEM_CD) AS CD");
                strSql.AppendLine("      , A.CON_ITEM_NM AS NM");
                strSql.AppendLine("      , ROW_NUMBER() OVER(ORDER BY CON_ITEM_NM) AS SEQ ");
                strSql.AppendLine("   FROM EQUIP_CONSUME_MGT A");
                strSql.AppendLine("  WHERE USE_YN = 'Y'");
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

            string sYmdFrom = DateEditFrom.EditValue.ToString().Replace("-", "").Substring(0, 8);
            string sYmdTo = DateEditTo.EditValue.ToString().Replace("-", "").Substring(0, 8);
            string sRdgbInOutGb = RdgbInOutgb.EditValue.ToString();
            string sItemNm = LkupMenuItemNm.EditValue?.ToString();

            StringBuilder strSql = new StringBuilder();

            strSql.Clear();
            strSql.AppendLine(" ");
            strSql.AppendLine(" SELECT A.ITEM_INOUT_GB AS IN_OUT_GB ");
            strSql.AppendLine(" 	 , A.OCCUR_DT AS OCCUR_DT ");
            strSql.AppendLine(" 	 , A.CON_ITEM_AMT AS ITEM_AMT ");
            strSql.AppendLine(" 	 , A.CON_ITEM_UNPR AS ITEM_UNPR ");
            strSql.AppendLine(" 	 , A.CON_ITEM_PRICE AS ITEM_PRICE ");
            strSql.AppendLine(" 	 , A.CON_SUPL_DEALER AS SUPL_DEALER ");
            strSql.AppendLine(" 	 , B.CON_ITEM_NM AS ITEM_NM ");
            strSql.AppendLine(" 	 , B.CON_ITEM_SPEC AS ITEM_SPEC ");
            strSql.AppendLine(" 	 , B.CON_ITEM_PURP AS ITEM_PURP "); 
            strSql.AppendLine(" 	 , A.CON_STORAGE AS STORAGE ");
            strSql.AppendLine("      , A.ENT_DT ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.ENT_ID AS NUMERIC) IS NULL THEN A.ENT_ID ELSE DBO.FN_USRNM(A.ENT_ID) END AS ENT_ID");
            strSql.AppendLine("      , A.MFY_DT                                                                                                ");
            strSql.AppendLine("      , CASE WHEN TRY_PARSE(A.MFY_ID AS NUMERIC) IS NULL THEN A.MFY_ID ELSE DBO.FN_USRNM(A.MFY_ID) END AS MFY_ID");
            strSql.AppendLine("   FROM EQUIP_CONSUME_HISTORY A  "); 
            strSql.AppendLine("   LEFT OUTER JOIN EQUIP_CONSUME_MGT B  ");
            strSql.AppendLine("     ON B.CON_ITEM_CD = A.CON_ITEM_CD  ");
            strSql.AppendLine("  WHERE A.OCCUR_DT >= '" + sYmdFrom + "' ");
            strSql.AppendLine("    AND A.OCCUR_DT <= '" + sYmdTo + "' ");
            if(! sRdgbInOutGb.Equals("A")) strSql.AppendLine("    AND A.ITEM_INOUT_GB = '" + sRdgbInOutGb + "' ");
            if(! string.IsNullOrEmpty(sItemNm)) strSql.AppendLine("    AND B.CON_ITEM_CD = '" + sItemNm + "' ");
            strSql.AppendLine("  ORDER BY A.OCCUR_DT");

            DataTable dt = DBConn.GetDataTable(DBConn.dbCon, strSql.ToString());
            GridRetr.DataSource = dt;
        }

        private void BtnExport_Click(object sender, EventArgs e)
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

        private void RdgbInOutgb_SelectedIndexChanged(object sender, EventArgs e)
        {
            BtnRetr_Click(null, null);
        }

        private void GridViewRetr_CustomColumnDisplayText(object sender, DevExpress.XtraGrid.Views.Base.CustomColumnDisplayTextEventArgs e)
        {
            if (e.Column.FieldName.Equals("OCCUR_DT"))
            {
                if(e.DisplayText.Length >= 10)
                {
                    return;
                }
                else if(e.DisplayText.Length == 8)
                {
                    string sYmd = e.DisplayText;
                    string sResult = sYmd.Substring(0, 4) + "-" + sYmd.Substring(4, 2) + "-" + sYmd.Substring(6, 2);
                    e.DisplayText = sResult;
                }
            }
        }

        private void DateEditFrom_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void DateEditTo_Leave(object sender, EventArgs e)
        {
            if (!ComLib.ClsFunc.ValidChkFromToRetrYmd(DateEditFrom, DateEditTo))
            {
                MessageBox.Show("조회 시작일자가 종료일자 보다 이후 일 수 없습니다.\r\n시작일자를 종료일자로 변환합니다.");
                DateEditFrom.EditValue = DateEditTo.EditValue;
                return;
            }
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void EquipConsumeInOutRpt_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                
            }
            else if (e.KeyCode == Keys.F1)
            {
                //BtnCrete_Click(null, null);
            }
            else if (e.KeyCode == Keys.F3)
            {
                //BtnSave_Click(null, null);
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
                BtnExport_Click(null, null);
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

        private void EquipConsumeInOutRpt_TextChanged(object sender, EventArgs e)
        {
            if (this.Text.Contains(FmMainToolBar2.SAVE_LAYOUT_LOADING_NAME))
            {
                string[] sArrText = this.Text.Split('_');
                ComnEtcFunc.SaveLayout(FmMainToolBar2.UserID, this.Name, arrGrdView);
                this.Text = sArrText[0];

                XtraMessageBox.Show(string.Format("{0}의 레이아웃이 성공적으로 저장되었습니다.", this.Text));
            }
        }

        private void LkupMenuItemNm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                BtnRetr.PerformClick();
        }
    }
}